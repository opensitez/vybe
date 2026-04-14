use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;
use std::sync::{Arc, Mutex, Weak as ArcWeak};

use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::{EventLoop, Task};
use crate::fiber::{Fiber, SavedFrame};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;
use crate::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation, Value};

const MAX_FRAMES: usize = 256;
const MAX_STACK: usize = 65536;

/// Result of VM execution — may complete or suspend for async.
pub enum ExecResult {
    /// Execution completed with a value.
    Done(Value),
    /// Execution suspended — waiting for a Promise to resolve.
    /// Contains the promise ID the fiber is waiting on.
    Suspended(u64),
}

/// Restricted context passed to host functions.
/// Provides only the capabilities a host function needs:
/// - Invoke VM callbacks (for LINQ, event handlers, etc.)
/// - Access linear memory (for WASI filesystem, network, etc.)
/// - Access user-defined host state (GUI queue, side effects, etc.)
///
/// Does NOT expose: globals, stack, frames, bytecode, type registry.
/// This matches the WASM security model (Wasmtime Caller<State>).
pub struct HostContext<'a> {
    /// Invoke a VM function reference with arguments.
    /// This is the ONLY way host functions can call back into the VM.
    invoker: Option<&'a mut dyn FnMut(&Value, &[Value]) -> Value>,
    /// Linear memory access (WASM MVP memory[0]).
    pub memory: Option<&'a mut [u8]>,
}

impl<'a> HostContext<'a> {
    /// Call a VM function reference from host code.
    /// Returns Value::Null if no invoker is available.
    pub fn invoke(&mut self, func_ref: &Value, args: &[Value]) -> Value {
        if let Some(ref mut invoker) = self.invoker {
            invoker(func_ref, args)
        } else {
            Value::Null
        }
    }

    /// Create an empty context (for host functions that don't need callbacks).
    pub fn empty() -> Self {
        HostContext { invoker: None, memory: None }
    }
}

/// Host function signature. Receives restricted context + args, returns a value.
/// Host function signature.
pub type HostFn = Arc<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>;

/// WASM import resolution target. An import can resolve to:
/// - A host function (provided by the embedder)
/// - A component-exported function (another module's code)
/// - A stdlib redirect (global function registered at runtime)
#[derive(Clone)]
pub enum ImportTarget {
    /// Index into VM::host_fns
    Host(usize),
    /// Chunk index + arity — calls a function defined in another component
    ChunkFn { chunk_index: usize, arity: u8 },
    /// Runtime global lookup (stdlib functions registered via globals)
    StdlibRedirect(String),
}

#[derive(Debug, Clone)]
struct CallFrame {
    chunk_index: usize,
    ip: usize,
    base: usize,
    upvalues: Vec<Arc<Mutex<Upvalue>>>,
}

/// Exception handler entry — pushed by try_start, popped by try_end or catch.
#[derive(Debug, Clone)]
struct ExceptionHandler {
    /// Instruction pointer to jump to on catch.
    catch_ip: usize,
    /// Chunk index the handler was registered in.
    _chunk_index: usize,
    /// Stack depth when try_start was executed (for unwinding).
    stack_depth: usize,
    /// Call frame depth when try_start was executed.
    frame_depth: usize,
    /// Exception tag index (0 = catch-all, N = typed catch for tag N).
    /// References chunk.exception_tags[tag] for the type name.
    tag: u8,
}

/// A language-agnostic bytecode virtual machine.
///
/// The VM has no built-in functions or language-specific semantics.
/// The host (compiler runtime) registers native functions via `register_host_fn`
/// and sets up globals before calling `run`.
pub struct VM {
    pub chunks: Vec<Chunk>,
    frames: Vec<CallFrame>,
    stack: Vec<Value>,
    pub globals: HashMap<String, Value>,
    open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    host_fns: Vec<HostFn>,
    /// Registry: (module, name) → index into host_fns.
    pub host_registry: HashMap<(String, String), usize>,
    /// Import resolution table: import_index → resolved target.
    /// WASM-aligned: imports can resolve to host functions OR component-exported functions.
    import_table: Vec<ImportTarget>,
    /// Exception handler stack (WASM exception proposal).
    exception_handlers: Vec<ExceptionHandler>,
    /// Event loop for async operations (shared with host functions).
    pub event_loop: Rc<RefCell<EventLoop>>,
    /// WASM GC-style type definitions with vtable method dispatch.
    pub type_registry: crate::typedef::TypeRegistry,
    /// Linear memory (WASM MVP) — byte buffer for binary data.
    /// This is memory index 0 for backward compatibility.
    pub memory: SharedMemory,
    /// Additional memories for multi-memory support.
    /// memory index 0 = self.memory, index 1+ = extra_memories[i-1].
    extra_memories: Vec<Vec<u8>>,
    /// Currently selected memory index (for load/store ops). Default 0.
    active_memory: usize,
    /// Function table (WASM MVP) — for call_indirect.
    pub func_table: Vec<Value>,
    /// Block label stack for structured control flow.
    label_stack: Vec<LabelEntry>,
    /// Callback invoker for host functions (cached allocation).
    callback_invoker: Option<Box<dyn FnMut(&Value, &[Value]) -> Value>>,
    /// When true, enforce strict WASM isolation:
    /// - Module-scoped globals (prefixed by module name)
    /// - Per-module memory (separate linear memory per component)
    /// Default false for trusted code (shared globals for cross-language interop).
    pub strict_isolation: bool,
    /// Module prefix for current execution context (used when strict_isolation=true).
    pub module_prefix: Option<String>,
    /// CLS case alias map: lowercase → canonical casing.
    /// When a global_get fails, tries lowercase lookup in this map to find
    /// the canonical-cased name. Enables cross-language name resolution
    /// (VB lowercase → C# PascalCase, COBOL UPPERCASE → VB lowercase).
    pub case_aliases: HashMap<String, String>,
    /// Finalizer registry: maps object identity to callback.
    /// When an object's strong count reaches the weak+finalizer threshold,
    /// the callback is queued for execution.
    finalizers: Vec<FinalizerEntry>,
    /// Active threads spawned by thread_spawn opcode.
    /// Maps thread_id → JoinHandle that returns the serialized result.
    thread_handles: HashMap<i32, std::thread::JoinHandle<Vec<u8>>>,
    /// Next thread ID to assign.
    next_thread_id: i32,
    /// Execution trace: when true, print every opcode + stack top.
    /// Enable via `vm.set_trace(true)` or `VYBE_TRACE=1` env var.
    trace: bool,
}

/// A registered finalizer for an object.
#[derive(Clone)]
struct FinalizerEntry {
    /// Weak reference to the target object.
    target: ArcWeak<Mutex<crate::value::Object>>,
    /// Callback to invoke when the object is about to be collected.
    callback: Value,
}

/// Entry in the structured control flow label stack.
#[derive(Debug, Clone)]
struct LabelEntry {
    /// Instruction offset to jump to on `br` (end of block, or start of loop).
    target: usize,
    /// True if this is a loop (continue jumps to start), false if block (break jumps to end).
    is_loop: bool,
}

impl VM {
    pub fn new() -> Self {
        VM {
            chunks: Vec::new(),
            frames: Vec::new(),
            stack: Vec::with_capacity(256),
            globals: HashMap::new(),
            open_upvalues: Vec::new(),
            host_fns: Vec::new(),
            host_registry: HashMap::new(),
            import_table: Vec::<ImportTarget>::new(),
            exception_handlers: Vec::new(),
            event_loop: Rc::new(RefCell::new(EventLoop::new())),
            type_registry: crate::typedef::TypeRegistry::new(),
            memory: SharedMemory::default(),
            extra_memories: Vec::new(),
            active_memory: 0,
            func_table: Vec::new(),
            label_stack: Vec::new(),
            callback_invoker: None,
            strict_isolation: false,
            module_prefix: None,
            case_aliases: HashMap::new(),
            finalizers: Vec::new(),
            thread_handles: HashMap::new(),
            next_thread_id: 1,
            trace: std::env::var("VYBE_TRACE").map_or(false, |v| v == "1" || v == "true"),
        }
    }

    /// Enable or disable execution tracing. When enabled, every opcode
    /// execution prints the chunk name, offset, opcode, and stack top.
    /// Can also be enabled via `VYBE_TRACE=1` environment variable.
    pub fn set_trace(&mut self, enabled: bool) {
        self.trace = enabled;
    }

    /// Capture the current call stack for error reporting.
    pub fn capture_call_stack(&self) -> Vec<crate::error::StackFrame> {
        self.frames.iter().rev().map(|f| {
            let chunk = &self.chunks[f.chunk_index];
            let line = chunk.get_line(f.ip.saturating_sub(1));
            crate::error::StackFrame {
                chunk_name: chunk.name.clone(),
                offset: f.ip,
                line,
            }
        }).collect()
    }

    /// Dump disassembled bytecode for all chunks. Useful for debugging
    /// test failures — call after `compile()` to see what was emitted.
    /// Returns a formatted string, one chunk per section.
    pub fn dump_bytecode(&self) -> String {
        let mut out = String::new();
        for (i, chunk) in self.chunks.iter().enumerate() {
            out.push_str(&format!("\n── Chunk {} ──\n", i));
            out.push_str(&crate::debug::disassemble(chunk));
        }
        out
    }

    /// Dump bytecode for a specific chunk by index.
    pub fn dump_chunk(&self, index: usize) -> String {
        if index < self.chunks.len() {
            crate::debug::disassemble(&self.chunks[index])
        } else {
            format!("Chunk {} not found (have {})", index, self.chunks.len())
        }
    }

    /// Check if native Vybe host functions are registered.
    pub fn has_vybe_host(&self) -> bool {
        self.host_registry.contains_key(&("vybe:array".to_string(), "redim".to_string()))
    }

    /// Evaluate a constant expression (Extended Const Expressions).
    /// Used for global initialization at load time.
    fn eval_const_expr(&self, expr: &crate::chunk::ConstExpr) -> Value {
        use crate::chunk::ConstExpr;
        match expr {
            ConstExpr::Value(v) => v.clone(),
            ConstExpr::GlobalGet(name) => {
                self.globals.get(name).cloned().unwrap_or(Value::Null)
            }
            ConstExpr::Add(left, right) => {
                let l = self.eval_const_expr(left);
                let r = self.eval_const_expr(right);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_add(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_add(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a + b),
                    _ => Value::F64(l.as_f64() + r.as_f64()),
                }
            }
            ConstExpr::Mul(left, right) => {
                let l = self.eval_const_expr(left);
                let r = self.eval_const_expr(right);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_mul(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_mul(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a * b),
                    _ => Value::F64(l.as_f64() * r.as_f64()),
                }
            }
            ConstExpr::RefFunc(chunk_idx) => {
                if *chunk_idx < self.chunks.len() {
                    let chunk = &self.chunks[*chunk_idx];
                    let func = crate::value::Function {
                        name: Some(chunk.name.clone()),
                        arity: chunk.arity,
                        chunk_index: *chunk_idx,
                        upvalues: Vec::new(),
                    };
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::Function(func);
                    Value::Object(Arc::new(Mutex::new(obj)))
                } else {
                    Value::Null
                }
            }
        }
    }

    /// Get the size (in bytes) of the currently active memory.
    fn active_mem_len(&self) -> usize {
        if self.active_memory == 0 {
            self.memory.len()
        } else {
            let idx = self.active_memory - 1;
            if idx < self.extra_memories.len() { self.extra_memories[idx].len() } else { 0 }
        }
    }

    /// Grow the currently active memory by `pages` pages. Returns old page count.
    fn active_mem_grow(&mut self, pages: usize) -> usize {
        if self.active_memory == 0 {
            self.memory.grow(pages)
        } else {
            let idx = self.active_memory - 1;
            if idx >= self.extra_memories.len() {
                self.extra_memories.resize_with(idx + 1, Vec::new);
            }
            let mem = &mut self.extra_memories[idx];
            let old_pages = mem.len() / 65536;
            mem.resize(mem.len() + pages * 65536, 0);
            old_pages
        }
    }

    /// Get a reference to a specific extra memory by index (index > 0 only).
    fn extra_mem(&self, idx: usize) -> &[u8] {
        if idx == 0 || idx - 1 >= self.extra_memories.len() {
            &[]
        } else {
            &self.extra_memories[idx - 1]
        }
    }

    /// Get a mutable reference to a specific extra memory by index (index > 0 only).
    fn extra_mem_mut(&mut self, idx: usize) -> &mut Vec<u8> {
        let i = idx - 1;
        if i >= self.extra_memories.len() {
            self.extra_memories.resize_with(i + 1, Vec::new);
        }
        &mut self.extra_memories[i]
    }

    /// Run any pending finalizers for objects whose strong count has dropped.
    /// Returns collected callbacks that should be invoked by the caller.
    pub fn collect_dead_finalizers(&mut self) -> Vec<Value> {
        let mut callbacks = Vec::new();
        let mut i = 0;
        while i < self.finalizers.len() {
            if self.finalizers[i].target.strong_count() == 0 {
                let entry = self.finalizers.remove(i);
                callbacks.push(entry.callback);
            } else {
                i += 1;
            }
        }
        callbacks
    }

    /// Register a host function with a (module, name) pair.
    /// Also adds it to the function table for call_indirect dispatch.
    pub fn register_host_fn(&mut self, module: &str, name: &str, f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>) {
        let idx = self.host_fns.len();
        self.host_fns.push(Arc::from(f));
        self.host_registry.insert((module.to_string(), name.to_string()), idx);
        // Add to function table — func_table index == host_fns index for host functions
        while self.func_table.len() <= idx {
            self.func_table.push(Value::Null);
        }
        // Store as a lightweight marker — call_indirect will recognize host fn indices
        let mut obj = Object::new();
        obj.kind = ObjectKind::HostFunction(idx);
        self.func_table[idx] = Value::Object(Arc::new(Mutex::new(obj)));
    }

    /// Create a HostContext with callback capability for host functions.
    fn make_host_context(&mut self) -> HostContext<'_> {
        // We can't pass &mut self into the closure directly due to borrow rules.
        // Instead, we pass raw pointers — this is safe because the HostContext
        // lifetime is strictly scoped within the host function call.
        let vm_ptr = self as *mut VM;
        HostContext {
            invoker: Some(unsafe {
                // SAFETY: vm_ptr is valid for the duration of the host function call.
                // The host function cannot outlive the call_import/call_value scope.
                let vm_ref: &mut VM = &mut *vm_ptr;
                vm_ref.get_invoker()
            }),
            memory: None, // TODO: pass memory when needed
        }
    }

    /// Get a mutable reference to the invoker closure.
    fn get_invoker(&mut self) -> &mut dyn FnMut(&Value, &[Value]) -> Value {
        // This is stored as a field to avoid repeated allocation
        if self.callback_invoker.is_none() {
            let vm_ptr = self as *mut VM;
            self.callback_invoker = Some(Box::new(move |func_ref: &Value, args: &[Value]| {
                // SAFETY: vm_ptr is valid during host function execution
                let vm = unsafe { &mut *vm_ptr };
                vm.invoke_callback(func_ref, args)
            }));
        }
        self.callback_invoker.as_mut().unwrap().as_mut()
    }

    /// Invoke a VM function reference from host code.
    /// This is the WASM-compliant callback mechanism: host functions
    /// can call exported/internal VM functions during execution.
    ///
    /// Usage from a host function:
    ///   let result = vm.invoke_callback(&predicate, &[element]);
    pub fn invoke_callback(&mut self, func_ref: &Value, args: &[Value]) -> Value {
        let saved_frame_depth = self.frames.len();

        // Push function ref + args onto stack
        self.stack.push(func_ref.clone());
        for arg in args {
            self.stack.push(arg.clone());
        }

        // Call the function (pushes a new frame)
        if self.call_value(args.len()).is_err() {
            return Value::Null;
        }

        // Execute until the callback frame returns
        match self.execute_until(saved_frame_depth + 1) {
            Ok(val) => val,
            Err(_) => {
                // On error, unwind
                while self.frames.len() > saved_frame_depth {
                    self.frames.pop();
                }
                Value::Null
            }
        }
    }

    /// Get a type_id by name from the TypeRegistry.
    pub fn get_type_id(&self, name: &str) -> usize {
        self.type_registry.get_id(name).unwrap_or(0)
    }


    /// Load chunks and execute the script chunk (first in the new set).
    /// Appends to existing chunks so cross-language calls work (functions reference chunk indices).
    /// Resolves the import table against registered host functions.
    /// Run linked components with WASM Component Model isolation.
    /// Each component gets its own global namespace (prefixed).
    /// Cross-component communication happens ONLY through declared exports/imports.
    /// Type metadata is shared read-only for inheritance.
    pub fn run_components(&mut self, link_result: &crate::component::LinkResult, components: &[crate::component::Component]) -> Result<Value, VMError> {
        // Load all chunks
        let base_offset = self.chunks.len();
        self.chunks.extend(link_result.chunks.clone());

        // Load shared type table (read-only cross-module)
        for ((_, type_name), typedef) in &link_result.type_exports {
            // Types are shared — they enable cross-language inheritance
            let _ = (type_name, typedef);
        }

        // Load type tables from all chunks
        for chunk in &link_result.chunks {
            if !chunk.types.is_empty() {
                self.type_registry.load_type_table(&chunk.types);
            }
        }

        // Run each component's script chunk with module isolation
        let saved_isolation = self.strict_isolation;
        let saved_prefix = self.module_prefix.clone();
        self.strict_isolation = true;

        for (i, comp) in components.iter().enumerate() {
            let _chunk_offset = link_result.component_offsets[i] + base_offset;

            // Set module prefix for global isolation
            self.module_prefix = Some(comp.name.clone());

            // Inject imported function references into this module's scope
            for (iface, func_name) in &comp.imports {
                let key = (iface.clone(), func_name.clone());
                if let Some(export_impl) = link_result.exports.get(&key) {
                    let func_val = match export_impl {
                        crate::component::ExportImpl::ChunkFn(ci) => {
                            let adjusted_ci = ci + base_offset;
                            let chunk = &self.chunks[adjusted_ci];
                            let func = crate::value::Function {
                                name: Some(func_name.clone()),
                                arity: chunk.arity,
                                chunk_index: adjusted_ci,
                                upvalues: Vec::new(),
                            };
                            let obj = crate::value::Object {
                                properties: std::collections::HashMap::new(),
                                kind: crate::value::ObjectKind::Function(func),
                                type_id: 0, fields: Vec::new(),
                            };
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                        crate::component::ExportImpl::HostFn(idx) => {
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::HostFunction(*idx);
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                    };
                    // Store in module-scoped globals
                    let global_key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.insert(global_key, func_val.clone());
                    // Also store without prefix so the module's code can find it
                    // (the module emits global_get "func_name", which gets prefixed by strict_isolation)
                    let unprefixed = func_name.to_lowercase();
                    self.globals.insert(format!("{}::{}", comp.name, unprefixed), func_val);
                }
            }

            // Also inject exported functions from OTHER modules that this module imports
            // by making them available under the importing module's prefix
            for other_comp in components {
                if other_comp.name == comp.name { continue; }
                for ((_, func_name), export_impl) in &other_comp.exports {
                    let func_val = match export_impl {
                        crate::component::ExportImpl::ChunkFn(ci) => {
                            let other_offset = link_result.component_offsets[
                                components.iter().position(|c| c.name == other_comp.name).unwrap()
                            ] + base_offset;
                            let adjusted_ci = ci + other_offset;
                            if adjusted_ci >= self.chunks.len() { continue; }
                            let chunk = &self.chunks[adjusted_ci];
                            let func = crate::value::Function {
                                name: Some(func_name.clone()),
                                arity: chunk.arity,
                                chunk_index: adjusted_ci,
                                upvalues: Vec::new(),
                            };
                            let obj = crate::value::Object {
                                properties: std::collections::HashMap::new(),
                                kind: crate::value::ObjectKind::Function(func),
                                type_id: 0, fields: Vec::new(),
                            };
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                        crate::component::ExportImpl::HostFn(idx) => {
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::HostFunction(*idx);
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                    };
                    // Available to this module via its prefix
                    let key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.entry(key).or_insert(func_val.clone());
                    // Also store class constructors so inheritance works
                    // (the child module does global_get "ClassName" which becomes "child::classname")
                    let ctor_key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.entry(ctor_key).or_insert(func_val);
                }
            }

            // Run this component's chunks through the standard run() path
            // which handles import resolution, type tables, and frame setup
            let comp_chunks = comp.chunks.clone();
            match self.run(comp_chunks) {
                Ok(_) => {}
                Err(e) => {
                    self.strict_isolation = saved_isolation;
                    self.module_prefix = saved_prefix;
                    return Err(e);
                }
            }
        }

        // Restore isolation state
        self.strict_isolation = saved_isolation;
        self.module_prefix = saved_prefix;

        Ok(Value::Null)
    }

    /// Run linked chunks with a pre-resolved import table from the Linker.
    /// Used for bootstrap: Linker resolves imports at link time, VM just loads them.
    pub fn run_linked(&mut self, chunks: Vec<Chunk>, resolved_imports: Vec<ImportTarget>) -> Result<Value, VMError> {
        if chunks.is_empty() {
            return Ok(Value::Null);
        }
        let script_idx = self.chunks.len();
        // Offset ref_func indices
        let mut adjusted = chunks;
        if script_idx > 0 {
            for chunk in &mut adjusted {
                let code = &mut chunk.code;
                let mut ip = 0;
                while ip < code.len() {
                    let op_byte = code[ip];
                    if let Some(op) = Op::from_byte(op_byte) {
                        match op {
                            Op::ref_func => {
                                if ip + 2 < code.len() {
                                    let old_idx = ((code[ip + 1] as u16) << 8) | (code[ip + 2] as u16);
                                    let new_idx = old_idx + script_idx as u16;
                                    code[ip + 1] = (new_idx >> 8) as u8;
                                    code[ip + 2] = (new_idx & 0xff) as u8;
                                }
                                ip += 3 + 1;
                                if ip - 1 < code.len() {
                                    let uv_count = code[ip - 1] as usize;
                                    ip += uv_count * 2;
                                }
                                continue;
                            }
                            _ => {}
                        }
                        ip += op.encoded_len();
                        match op {
                            Op::call_import => { ip += 3; }
                            Op::br | Op::br_if_true | Op::br_if_false | Op::br_if_null
                            | Op::r#loop => { ip += 2; }
                            Op::try_start => { ip += 4; }
                            _ => {}
                        }
                    } else if op_byte == 0xFE && ip + 1 < code.len() {
                        ip += 2;
                    } else {
                        ip += 1;
                    }
                }
            }
        }
        self.chunks.extend(adjusted);

        // Use pre-resolved import table, adjusting ChunkFn indices by offset
        self.import_table.clear();
        for target in resolved_imports {
            match target {
                ImportTarget::ChunkFn { chunk_index, arity } => {
                    self.import_table.push(ImportTarget::ChunkFn {
                        chunk_index: chunk_index + script_idx,
                        arity,
                    });
                }
                ImportTarget::Host(idx) => {
                    self.import_table.push(ImportTarget::Host(idx));
                }
                ImportTarget::StdlibRedirect(name) => {
                    // Try to resolve against host registry first
                    let key_candidates = [
                        ("wasi:cli".to_string(), name.clone()),
                        ("vybe:console".to_string(), name.clone()),
                    ];
                    let mut resolved = false;
                    for key in &key_candidates {
                        if let Some(&idx) = self.host_registry.get(key) {
                            self.import_table.push(ImportTarget::Host(idx));
                            resolved = true;
                            break;
                        }
                    }
                    if !resolved {
                        self.import_table.push(ImportTarget::StdlibRedirect(name));
                    }
                }
            }
        }

        // Load type table
        {
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                let adjusted_types: Vec<_> = types.iter().map(|t| {
                    let mut entry = t.clone();
                    entry.methods = t.methods.iter().map(|(name, idx)| {
                        (name.clone(), idx + script_idx)
                    }).collect();
                    if let Some(ci) = entry.constructor_chunk {
                        entry.constructor_chunk = Some(ci + script_idx);
                    }
                    entry
                }).collect();
                self.type_registry.load_type_table(&adjusted_types);
            }
        }

        // Execute
        self.frames.push(CallFrame {
            chunk_index: script_idx,
            ip: 0,
            base: self.stack.len(),
            upvalues: Vec::new(),
        });
        self.stack.resize(self.stack.len() + self.chunks[script_idx].local_count as usize, Value::Null);
        self.execute()
    }

    pub fn run(&mut self, chunks: Vec<Chunk>) -> Result<Value, VMError> {
        if chunks.is_empty() {
            return Ok(Value::Null);
        }
        let script_idx = self.chunks.len(); // offset for new chunks
        // Offset ref_func indices in the new chunks so they point to correct positions
        let mut adjusted = chunks;
        if script_idx > 0 {
            for chunk in &mut adjusted {
                let code = &mut chunk.code;
                let mut ip = 0;
                while ip < code.len() {
                    let op_byte = code[ip];
                    if let Some(op) = Op::from_byte(op_byte) {
                        match op {
                            Op::ref_func => {
                                // ref_func has u16 chunk_index operand
                                if ip + 2 < code.len() {
                                    let old_idx = ((code[ip + 1] as u16) << 8) | (code[ip + 2] as u16);
                                    let new_idx = old_idx + script_idx as u16;
                                    code[ip + 1] = (new_idx >> 8) as u8;
                                    code[ip + 2] = (new_idx & 0xff) as u8;
                                }
                                ip += 3 + 1; // op + u16 + upvalue_count byte
                                // Skip upvalue descriptors
                                if ip - 1 < code.len() {
                                    let uv_count = code[ip - 1] as usize;
                                    ip += uv_count * 2;
                                }
                                continue;
                            }
                            _ => {}
                        }
                        ip += op.encoded_len();
                        // Skip operands based on opcode
                        match op {
                            Op::r#const | Op::local_get | Op::local_set | Op::global_get
                            | Op::global_set | Op::struct_get | Op::struct_set | Op::struct_new
                            | Op::array_new | Op::ref_test => { ip += 2; }
                            Op::call | Op::call_ref | Op::upvalue_get | Op::upvalue_set
                            | Op::str_concat_n => { ip += 1; }
                            Op::call_import => { ip += 3; } // u16 + u8
                            Op::br | Op::br_if_true | Op::br_if_false | Op::br_if_null
                            | Op::r#loop => { ip += 2; }
                            Op::try_start => { ip += 4; }
                            _ => {}
                        }
                    } else if op_byte == 0xFE && ip + 1 < code.len() {
                        ip += 2; // extended opcode
                    } else {
                        ip += 1;
                    }
                }
            }
        }
        self.chunks.extend(adjusted);

        // Resolve imports for ALL new chunks (not just script chunk).
        // Each chunk has its own import list. We build one unified import table
        // by scanning all chunks and mapping their import indices to host functions.
        // The trick: all chunks compiled by the same compiler share the same import list
        // (imports are added to chunks[0] by all compilers). For multi-module programs,
        // different modules may have different imports. We resolve the union.
        self.import_table.clear();
        for (_i, import) in self.chunks[script_idx].imports.iter().enumerate() {
            // 1. Try host function registry (exact module:name match)
            let key = (import.module.clone(), import.name.clone());
            if let Some(&idx) = self.host_registry.get(&key) {
                self.import_table.push(ImportTarget::Host(idx));
                continue;
            }
            // 2. Wildcard module "*" — resolve from globals (cross-language or same-language)
            if import.module == "*" {
                // Check lowercase and original case
                let candidates = [
                    import.name.clone(),
                    import.name.to_lowercase(),
                ];
                let found = candidates.iter().find(|g| self.globals.contains_key(g.as_str()));
                if let Some(global_name) = found {
                    self.import_table.push(ImportTarget::StdlibRedirect(global_name.clone()));
                    continue;
                }
            }
            // 3. Check for stdlib global
            let candidates = [
                format!("__vybe_{}", import.name),
                format!("__vybe_{}", import.name.to_lowercase()),
            ];
            let found = candidates.iter().find(|g| self.globals.contains_key(g.as_str()));
            if let Some(global_name) = found {
                self.import_table.push(ImportTarget::StdlibRedirect(global_name.clone()));
            } else {
                return Err(VMError::new(format!(
                    "Unresolved import: \"{}\" \"{}\"", import.module, import.name
                )));
            }
        }

        // Load type table from the script chunk (WASM GC type section).
        // Registers user-defined class types and their vtable methods.
        // Sets __tid_<name> globals so constructors can stamp type_id.
        {
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                // Adjust chunk indices in methods (same offset as ref_func)
                let adjusted_types: Vec<_> = types.iter().map(|t| {
                    let mut entry = t.clone();
                    entry.methods = t.methods.iter().map(|(name, idx)| {
                        (name.clone(), idx + script_idx)
                    }).collect();
                    entry
                }).collect();
                self.type_registry.load_type_table(&adjusted_types);
                // Set __tid_<name> globals for each registered type
                for entry in &adjusted_types {
                    if let Some(tid) = self.type_registry.get_id(&entry.name) {
                        let key = format!("__tid_{}", entry.name.to_lowercase());
                        self.globals.insert(key, Value::I32(tid as i32));
                    }
                }
            }
        }

        // Evaluate global initializers (Extended Const Expressions).
        // These are computed at load time before any code runs.
        //
        // Polyfill rule: if the host has already populated a global before
        // run() is called (e.g. Vybe overriding `__vybe_pow` with native f64
        // pow), do NOT clobber it with the bundled stdlib fallback. The
        // stdlib chunks installed via global_inits are the portable
        // implementation that runs on standard WASM VMs; on Vybe they yield
        // to the optimized native version.
        {
            let inits = self.chunks[script_idx].global_inits.clone();
            for gi in &inits {
                if matches!(self.globals.get(&gi.name), None | Some(Value::Null)) {
                    let val = self.eval_const_expr(&gi.init);
                    self.globals.insert(gi.name.clone(), val);
                }
            }
        }

        self.frames.push(CallFrame {
            chunk_index: script_idx,
            ip: 0,
            base: 0,
            upvalues: Vec::new(),
        });

        let local_count = self.chunks[script_idx].local_count as usize;
        for _ in 0..local_count {
            self.stack.push(Value::Null);
        }

        // Run synchronous code
        let result = self.execute_with_async()?;

        // Event loop — process async tasks until all done
        match result {
            ExecResult::Done(val) => {
                self.run_event_loop()?;
                Ok(val)
            }
            ExecResult::Suspended(_) => {
                self.run_event_loop()?;
                Ok(Value::Null)
            }
        }
    }

    /// Run the event loop until all pending tasks are processed.
    fn run_event_loop(&mut self) -> Result<(), VMError> {
        loop {
            let has_pending = self.event_loop.borrow().has_pending();
            if !has_pending { break; }

            // 1. Drain all microtasks
            let microtasks: Vec<Task> = {
                let mut el = self.event_loop.borrow_mut();
                let mut tasks = Vec::new();
                while let Some(task) = el.next_microtask() {
                    tasks.push(task);
                }
                tasks
            };

            for task in microtasks {
                match task {
                    Task::Microtask { callback, value } => {
                        self.invoke(&callback, &[value])?;
                    }
                    Task::ResumeFiber(fiber) => {
                        self.resume_fiber(fiber)?;
                    }
                    _ => {}
                }
            }

            // 2. Wait for and process one macrotask (timer)
            {
                let el = self.event_loop.borrow();
                el.wait_for_next();
            }
            let timer = self.event_loop.borrow_mut().next_ready_timer();
            if let Some(Task::Timer { callback, .. }) = timer {
                self.invoke(&callback, &[])?;
            }
        }
        Ok(())
    }

    /// Resume a suspended fiber — restore its state and continue execution.
    fn resume_fiber(&mut self, fiber: Fiber) -> Result<Value, VMError> {
        // Restore state from fiber
        self.stack = fiber.stack;
        self.frames = fiber.frames.into_iter().map(|f| CallFrame {
            chunk_index: f.chunk_index,
            ip: f.ip,
            base: f.base,
            upvalues: f.upvalues,
        }).collect();
        self.open_upvalues = fiber.open_upvalues;

        // Push the resolved value onto the stack (this is what `await` returns)
        if let Some(val) = fiber.resume_value {
            self.push(val)?;
        }

        // Continue execution
        match self.execute_with_async()? {
            ExecResult::Done(val) => Ok(val),
            ExecResult::Suspended(_) => Ok(Value::Null), // re-suspended, event loop will handle
        }
    }

    /// JSPI: Resolve a suspended promise and resume execution.
    /// Called by the runtime/event loop when an async operation completes.
    /// `promise_id` identifies which suspension to resume.
    /// `value` is the resolved value that becomes the return of the host call.
    pub fn jspi_resolve(&mut self, promise_id: u64, value: Value) -> Result<Value, VMError> {
        let fiber = self.event_loop.borrow_mut().resolve_promise(promise_id, value);
        if let Some(fiber) = fiber {
            self.resume_fiber(fiber)
        } else {
            Ok(Value::Null)
        }
    }

    /// Check if there are any JSPI-suspended fibers waiting for resolution.
    pub fn has_pending_jspi(&self) -> bool {
        self.event_loop.borrow().has_pending()
    }

    /// Save the current execution state to a Fiber.
    fn save_fiber(&mut self) -> Fiber {
        let frames = self.frames.drain(..).map(|f| SavedFrame {
            chunk_index: f.chunk_index,
            ip: f.ip,
            base: f.base,
            upvalues: f.upvalues,
        }).collect();
        let stack = self.stack.drain(..).collect();
        let upvalues = self.open_upvalues.drain(..).collect();
        Fiber::new(stack, frames, upvalues)
    }

    /// Call a function value after the initial run() has completed.
    /// Used for event callbacks — the VM state (globals, chunks) is preserved.
    pub fn invoke(&mut self, callee: &Value, args: &[Value]) -> Result<Value, VMError> {
        // Close any remaining open upvalues before clearing the stack,
        // so closures retain their captured values.
        self.close_upvalues(0);
        self.stack.clear();
        self.frames.clear();

        self.push(callee.clone())?;
        for arg in args {
            self.push(arg.clone())?;
        }

        self.call_value(args.len())?;

        // HostFunction calls are handled inline by call_value (no frame pushed).
        // If no frame was pushed, the result is already on the stack.
        if self.frames.is_empty() {
            return Ok(self.stack.pop().unwrap_or(Value::Null));
        }

        self.execute()
    }

    // -- Stack --

    fn push(&mut self, value: Value) -> Result<(), VMError> {
        if self.stack.len() >= MAX_STACK {
            return Err(VMError::new("Stack overflow"));
        }
        self.stack.push(value);
        Ok(())
    }

    fn pop(&mut self) -> Value {
        self.stack.pop().expect("stack underflow")
    }

    fn peek(&self, distance: usize) -> &Value {
        &self.stack[self.stack.len() - 1 - distance]
    }

    // -- Frame --

    fn frame(&self) -> &CallFrame {
        self.frames.last().expect("no frame")
    }

    fn frame_mut(&mut self) -> &mut CallFrame {
        self.frames.last_mut().expect("no frame")
    }

    fn read_byte(&mut self) -> u8 {
        let f = self.frame();
        let byte = self.chunks[f.chunk_index].code[f.ip];
        self.frame_mut().ip += 1;
        byte
    }

    fn read_u16(&mut self) -> u16 {
        let hi = self.read_byte() as u16;
        let lo = self.read_byte() as u16;
        (hi << 8) | lo
    }

    fn read_i16(&mut self) -> i16 {
        self.read_u16() as i16
    }

    fn get_constant(&self, index: u16) -> Value {
        let f = self.frame();
        self.chunks[f.chunk_index].constants[index as usize].clone()
    }

    fn constant_str(&self, index: u16) -> String {
        match &self.get_constant(index) {
            Value::String(s) => s.to_string(),
            v => format!("{}", v),
        }
    }

    // -- SIMD helpers --
    fn simd_i32x4_binop(&mut self, f: impl Fn(i32, i32) -> i32) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = i32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                let lb = i32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                out[i*4..i*4+4].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    fn simd_f64x2_binop(&mut self, f: impl Fn(f64, f64) -> f64) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                let lb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                out[i*8..i*8+8].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    fn simd_f64x2_cmp(&mut self, f: impl Fn(f64, f64) -> bool) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                let lb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                let mask: u64 = if f(la, lb) { u64::MAX } else { 0 };
                out[i*8..i*8+8].copy_from_slice(&mask.to_le_bytes());
            }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    fn simd_f32x4_binop(&mut self, f: impl Fn(f32, f32) -> f32) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                let lb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                out[i*4..i*4+4].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    fn simd_i8x16_binop(&mut self, f: impl Fn(u8, u8) -> u8) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..16 { out[i] = f(va[i], vb[i]); }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    fn simd_i16x8_binop(&mut self, f: impl Fn(i16, i16) -> i16) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..8 {
                let la = i16::from_le_bytes([va[i*2], va[i*2+1]]);
                let lb = i16::from_le_bytes([vb[i*2], vb[i*2+1]]);
                out[i*2..i*2+2].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }

    /// Test if a value matches a type name (used by ref_test, ref_cast, br_on_cast).
    /// Supports: WASM GC type_id lookup, __type string matching, __types array
    /// (JS class inheritance chain), and __control_type for GUI controls.
    fn test_type(&self, val: &Value, target_name: &str) -> bool {
        match val {
            Value::Object(o) => {
                let ob = o.lock().unwrap();
                // Fast path: type_id is set (properly typed object)
                if ob.type_id > 0 {
                    if let Some(target_id) = self.type_registry.get_id(target_name) {
                        return self.type_registry.is_subtype(ob.type_id, target_id);
                    }
                    return false;
                }

                // Slow path: type_id == 0 — check __type / __control_type strings
                let obj_type = ob.properties.get("__type")
                    .map(|v| format!("{}", v).to_lowercase())
                    .or_else(|| ob.properties.get("__control_type")
                        .map(|v| format!("{}", v).to_lowercase()))
                    .unwrap_or_default();

                // Direct name match
                if obj_type == target_name { return true; }

                // Check via type registry (subtype relationship)
                if let Some(tid) = self.type_registry.get_id(&obj_type) {
                    if let Some(target_id) = self.type_registry.get_id(target_name) {
                        if self.type_registry.is_subtype(tid, target_id) {
                            return true;
                        }
                    }
                }

                // Check __types array (JS class inheritance chain)
                if let Some(Value::Object(types)) = ob.properties.get("__types") {
                    let t = types.lock().unwrap();
                    if let crate::value::ObjectKind::Array(ref elems) = t.kind {
                        let target_lower = target_name.to_lowercase();
                        if elems.iter().any(|e| format!("{}", e).to_lowercase() == target_lower) {
                            return true;
                        }
                    }
                }

                // Universal: everything is an "object"
                target_name == "object"
            }
            Value::String(_) => target_name == "string" || target_name == "object",
            Value::F64(_) | Value::I32(_) | Value::I64(_) => {
                target_name == "integer" || target_name == "double" || target_name == "number" || target_name == "object"
            }
            Value::Bool(_) => target_name == "boolean" || target_name == "object",
            Value::V128(_) => target_name == "v128",
            Value::WeakRef(weak) => {
                if let Some(strong) = weak.upgrade() {
                    self.test_type(&Value::Object(strong), target_name)
                } else {
                    false
                }
            }
            Value::Null | Value::Undefined => false,
        }
    }

    /// Check if an exception value matches a tag name.
    /// Works for: string exceptions (by content), objects with __type or __exception_type,
    /// and cross-language name matching (e.g., "ValueError", "TypeError").
    fn exception_value_matches(&self, val: &Value, tag_name: &str) -> bool {
        let tag_lower = tag_name.to_lowercase();
        match val {
            Value::String(s) => {
                // String exceptions: match if the string contains the tag name
                // e.g., throw "ValueError: invalid input" matches tag "ValueError"
                let s_lower = s.to_lowercase();
                s_lower.starts_with(&tag_lower) || s_lower.contains(&tag_lower)
            }
            Value::Object(o) => {
                let ob = o.lock().unwrap();
                // Check __exception_type property (set by language-specific throw)
                if let Some(et) = ob.properties.get("__exception_type") {
                    let et_str = format!("{}", et).to_lowercase();
                    if et_str == tag_lower { return true; }
                }
                // Check __type property
                if let Some(t) = ob.properties.get("__type") {
                    let t_str = format!("{}", t).to_lowercase();
                    if t_str == tag_lower { return true; }
                }
                // Check "name" property (JS Error convention)
                if let Some(n) = ob.properties.get("name") {
                    let n_str = format!("{}", n).to_lowercase();
                    if n_str == tag_lower { return true; }
                }
                // Check "message" property as fallback
                if let Some(m) = ob.properties.get("message") {
                    let m_str = format!("{}", m).to_lowercase();
                    if m_str.starts_with(&tag_lower) { return true; }
                }
                false
            }
            _ => false,
        }
    }

    // -- Execute --

    fn execute_with_async(&mut self) -> Result<ExecResult, VMError> {
        match self.execute() {
            Ok(val) => Ok(ExecResult::Done(val)),
            Err(e) if e.message.starts_with("__await__:") => {
                // Await suspension — extract promise ID
                let id: u64 = e.message["__await__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended(id))
            }
            Err(e) => Err(e),
        }
    }

    fn execute(&mut self) -> Result<Value, VMError> {
        self.execute_until(0)
    }

    /// Execute bytecode until frame depth drops to `min_depth`.
    /// `min_depth = 0` runs until halt (normal execution).
    /// `min_depth > 0` runs until a callback returns (for invoke_callback).
    fn execute_until(&mut self, min_depth: usize) -> Result<Value, VMError> {
        loop {
            let f = self.frame();
            let chunk = &self.chunks[f.chunk_index];

            if f.ip >= chunk.code.len() {
                if self.frames.len() <= 1.max(min_depth + 1) {
                    return Ok(self.stack.pop().unwrap_or(Value::Null));
                }
                let base = self.frame().base;
                self.frames.pop();
                self.stack.truncate(base);
                self.push(Value::Null)?;
                continue;
            }

            let byte = self.read_byte();
            let op = if byte == 0xFE {
                // Extended opcode: prefix 0xFE + extension byte
                let ext = self.read_byte();
                match Op::from_two_bytes(byte, ext) {
                    Some(op) => op,
                    None => return Err(VMError::new(format!("Invalid extended opcode: 0xFE 0x{:02X}", ext))),
                }
            } else {
                match Op::from_byte(byte) {
                    Some(op) => op,
                    None => return Err(VMError::new(format!("Invalid opcode: {}", byte))),
                }
            };

            // ── Execution trace ──────────────────────────────────────────
            if self.trace {
                let f = self.frame();
                let chunk_name = &self.chunks[f.chunk_index].name;
                let ip = f.ip;
                let stack_top = if self.stack.is_empty() {
                    "[]".to_string()
                } else {
                    let top = &self.stack[self.stack.len() - 1];
                    let depth = self.stack.len();
                    format!("[{}] (depth={})", top, depth)
                };
                eprintln!("  TRACE {:>12} @{:04} {:?}  stack: {}",
                    chunk_name, ip.saturating_sub(1), op, stack_top);
            }

            match op {
                Op::halt => {
                    if self.frames.len() <= 1 {
                        // Top-level halt: terminate execution
                        self.close_upvalues(0);
                        return Ok(if self.stack.is_empty() { Value::Null } else { self.pop() });
                    } else {
                        // Nested halt (e.g. script chunk called via bootstrap):
                        // act like return — pop frame and return null
                        let base = self.frame().base;
                        self.close_upvalues(base);
                        self.frames.pop();
                        self.stack.truncate(base);
                        self.push(Value::Null)?;
                    }
                }
                Op::unreachable => {
                    return Err(VMError::new("trap: unreachable executed"));
                }

                Op::r#const => {
                    let idx = self.read_u16();
                    let val = self.get_constant(idx);
                    self.push(val)?;
                }
                Op::drop => { self.pop(); }
                Op::dup => {
                    let val = self.peek(0).clone();
                    self.push(val)?;
                }

                // -- Variables --
                Op::local_get => {
                    let slot = self.read_u16() as usize;
                    let base = self.frame().base;
                    let val = self.stack[base + slot].clone();
                    self.push(val)?;
                }
                Op::local_set => {
                    let slot = self.read_u16() as usize;
                    let val = self.peek(0).clone();
                    let base = self.frame().base;
                    self.stack[base + slot] = val;
                }
                Op::global_get => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    // In strict isolation mode, prefix globals with module name
                    // to prevent cross-module access
                    let key = if self.strict_isolation {
                        if let Some(ref prefix) = self.module_prefix {
                            let prefixed = format!("{}::{}", prefix, name);
                            // Try prefixed first, then unprefixed (for exports)
                            if self.globals.contains_key(&prefixed) {
                                prefixed
                            } else {
                                name
                            }
                        } else { name }
                    } else { name };
                    let val = self.globals.get(&key).cloned().unwrap_or(Value::Undefined);
                    self.push(val)?;
                }
                Op::global_set => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let key = if self.strict_isolation {
                        if let Some(ref prefix) = self.module_prefix {
                            format!("{}::{}", prefix, name)
                        } else { name }
                    } else { name };
                    let val = self.peek(0).clone();
                    self.globals.insert(key, val);
                }
                Op::upvalue_get => {
                    let idx = self.read_byte() as usize;
                    let uv = self.frame().upvalues[idx].clone();
                    let val = match &uv.lock().unwrap().location {
                        UpvalueLocation::Open(si) => self.stack[*si].clone(),
                        UpvalueLocation::Closed(v) => v.clone(),
                    };
                    self.push(val)?;
                }
                Op::upvalue_set => {
                    let idx = self.read_byte() as usize;
                    let val = self.peek(0).clone();
                    let uv = self.frame().upvalues[idx].clone();
                    let mut u = uv.lock().unwrap();
                    match &mut u.location {
                        UpvalueLocation::Open(si) => self.stack[*si] = val,
                        UpvalueLocation::Closed(v) => *v = val,
                    }
                }

                // -- Properties --
                Op::struct_get => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let obj = self.pop();
                    // Auto-join thread when accessing .result on a Task/Thread object
                    if let Value::Object(ref o) = obj {
                        let needs_join = {
                            let o_ref = o.lock().unwrap();
                            (name == "result" || name == "exitcode")
                                && o_ref.properties.contains_key("__thread_id")
                                && !o_ref.properties.get("iscompleted").map(|v| v.as_bool()).unwrap_or(true)
                        };
                        if needs_join {
                            let tid = o.lock().unwrap().properties.get("__thread_id")
                                .map(|v| v.as_f64() as i32).unwrap_or(-1);
                            if let Some(handle) = self.thread_handles.remove(&tid) {
                                let _ = handle.join();
                                // Task object was updated by child thread
                            }
                        }
                        // Check for getter
                        let getter_key = format!("__get_{}", name);
                        let getter = o.lock().unwrap().properties.get(&getter_key).cloned();
                        if let Some(getter_fn) = getter {
                            self.push(getter_fn)?;
                            self.push(obj)?;
                            self.call_value(1)?;
                            continue;
                        }
                    }
                    self.push(self.resolve_property(&obj, &name)?)?;
                }
                Op::struct_set => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.pop();
                    let obj = self.pop();
                    if let Value::Object(o) = &obj {
                        // Check for setter: __set_{name}
                        let setter_key = format!("__set_{}", name);
                        let setter = o.lock().unwrap().properties.get(&setter_key).cloned();
                        if let Some(setter_fn) = setter {
                            // Call the setter synchronously. Save stack depth
                            // and restore after — invoke_callback leaks the
                            // return value and intermediate locals on the stack.
                            let stack_save = self.stack.len();
                            let _result = self.invoke_callback(&setter_fn, &[obj.clone(), val.clone()]);
                            self.stack.truncate(stack_save);
                            self.push(val)?;
                        } else {
                            o.lock().unwrap().set(name.clone(), val.clone());
                            self.push(val)?;
                        }
                    } else {
                        self.push(val)?;
                    }
                }
                Op::array_get => {
                    let key = self.pop();
                    let obj = self.pop();
                    match &obj {
                        Value::Object(o) => {
                            // Handle negative indices: x[-1] → x[len-1]
                            let k = {
                                let idx = key.as_f64() as i64;
                                if idx < 0 {
                                    let ob = o.lock().unwrap();
                                    let len = match &ob.kind {
                                        ObjectKind::Array(a) => a.len() as i64,
                                        _ => 0,
                                    };
                                    format!("{}", (len + idx).max(0))
                                } else {
                                    format!("{}", key)
                                }
                            };
                            let val = o.lock().unwrap().get(&k);
                            // If not found and object has __getitem__, call it
                            if matches!(val, Value::Null) {
                                let getitem = o.lock().unwrap().properties.get("__getitem__").cloned();
                                if let Some(func) = getitem {
                                    self.push(func)?;
                                    self.push(obj.clone())?; // self
                                    self.push(key)?; // key arg
                                    self.call_value(2)?;
                                    continue;
                                }
                            }
                            self.push(val)?;
                        }
                        Value::String(s) => {
                            let i = key.as_f64() as usize;
                            if let Some(ch) = s.chars().nth(i) {
                                self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                            } else {
                                self.push(Value::Null)?;
                            }
                        }
                        _ => self.push(Value::Null)?,
                    }
                }
                Op::array_set => {
                    let val = self.pop();
                    let key = self.pop();
                    let obj = self.pop();
                    if let Value::Object(o) = &obj {
                        // Check for __setitem__ dunder
                        let setitem = o.lock().unwrap().properties.get("__setitem__").cloned();
                        if let Some(func) = setitem {
                            self.push(func)?;
                            self.push(obj.clone())?; // self
                            self.push(key)?;          // key
                            self.push(val.clone())?;  // value
                            self.call_value(3)?;
                            self.pop(); // discard __setitem__ return
                            self.push(val)?;
                            continue;
                        }
                        let k = format!("{}", key);
                        o.lock().unwrap().set(k, val.clone());
                    }
                    self.push(val)?;
                }

                // -- Float arithmetic --
                Op::f64_add => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a + b))?;
                }
                Op::f64_sub => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a - b))?;
                }
                Op::f64_mul => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a * b))?;
                }
                Op::f64_div => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a / b))?;
                }
                Op::f64_mod => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a % b))?;
                }
                Op::f64_neg => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(-a))?;
                }

                // -- Integer arithmetic --
                Op::i32_add => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_add(b)))?;
                }
                Op::i32_sub => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_sub(b)))?;
                }
                Op::i32_mul => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_mul(b)))?;
                }
                Op::i32_div_s => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32(a.wrapping_div(b)))?;
                }
                Op::i32_div_u => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32((a / b) as i32))?;
                }
                Op::i32_rem_s => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32(a.wrapping_rem(b)))?;
                }
                Op::i32_rem_u => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32((a % b) as i32))?;
                }

                // -- i64 arithmetic --
                Op::i64_add => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_add(b)))?; }
                Op::i64_sub => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_sub(b)))?; }
                Op::i64_mul => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_mul(b)))?; }
                Op::i64_div_s => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64(a.wrapping_div(b)))?; }
                Op::i64_div_u => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64((a / b) as i64))?; }
                Op::i64_rem_s => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64(a.wrapping_rem(b)))?; }
                Op::i64_rem_u => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64((a % b) as i64))?; }
                Op::i64_and => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a & b))?; }
                Op::i64_or  => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a | b))?; }
                Op::i64_xor => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a ^ b))?; }
                Op::i64_shl   => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a << (b & 0x3f)))?; }
                Op::i64_shr_s => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a >> (b & 0x3f)))?; }
                Op::i64_shr_u => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64((a >> (b & 0x3f)) as i64))?; }
                Op::i64_rotl => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(a.rotate_left((b & 0x3f) as u32) as i64))?; }
                Op::i64_rotr => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(a.rotate_right((b & 0x3f) as u32) as i64))?; }
                Op::i64_clz => { let a = self.pop().as_i64(); self.push(Value::I64(a.leading_zeros() as i64))?; }
                Op::i64_ctz => { let a = self.pop().as_i64(); self.push(Value::I64(a.trailing_zeros() as i64))?; }
                Op::i64_popcnt => { let a = self.pop().as_i64(); self.push(Value::I64(a.count_ones() as i64))?; }

                // -- f64 math --
                Op::f64_abs => { let a = self.pop().as_f64(); self.push(Value::F64(a.abs()))?; }
                Op::f64_ceil => { let a = self.pop().as_f64(); self.push(Value::F64(a.ceil()))?; }
                Op::f64_floor => { let a = self.pop().as_f64(); self.push(Value::F64(a.floor()))?; }
                Op::f64_trunc => { let a = self.pop().as_f64(); self.push(Value::F64(a.trunc()))?; }
                Op::f64_nearest => { let a = self.pop().as_f64(); self.push(Value::F64(a.round()))?; }
                Op::f64_sqrt => { let a = self.pop().as_f64(); self.push(Value::F64(a.sqrt()))?; }
                Op::f64_min => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.min(b)))?; }
                Op::f64_max => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.max(b)))?; }
                Op::f64_copysign => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.copysign(b)))?; }

                // -- f32 (promoted to f64) --
                Op::f32_abs => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).abs() as f64))?; }
                Op::f32_neg => { let a = self.pop().as_f64(); self.push(Value::F64(-(a as f32) as f64))?; }
                Op::f32_ceil => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).ceil() as f64))?; }
                Op::f32_floor => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).floor() as f64))?; }
                Op::f32_trunc => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).trunc() as f64))?; }
                Op::f32_nearest => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).round() as f64))?; }
                Op::f32_sqrt => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).sqrt() as f64))?; }
                Op::f32_min => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).min(b as f32) as f64))?; }
                Op::f32_max => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).max(b as f32) as f64))?; }
                Op::f32_copysign => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).copysign(b as f32) as f64))?; }

                // -- WASM select --
                Op::select => {
                    let cond = self.pop().as_i32();
                    let val2 = self.pop();
                    let val1 = self.pop();
                    self.push(if cond != 0 { val1 } else { val2 })?;
                }

                // -- i32 rotation and bit counting --
                Op::i32_rotl => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32(a.rotate_left(b & 0x1f) as i32))?; }
                Op::i32_rotr => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32(a.rotate_right(b & 0x1f) as i32))?; }
                Op::i32_clz => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.leading_zeros() as i32))?; }
                Op::i32_ctz => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.trailing_zeros() as i32))?; }
                Op::i32_popcnt => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.count_ones() as i32))?; }

                // -- eqz --
                Op::i32_eqz => { let a = self.pop().as_i32(); self.push(Value::Bool(a == 0))?; }
                Op::i64_eqz => { let a = self.pop().as_i64(); self.push(Value::Bool(a == 0))?; }

                // -- String --
                Op::str_concat => {
                    let b = self.pop();
                    let a = self.pop();
                    let s = format!("{}{}", a, b);
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                Op::str_concat_n => {
                    let count = self.read_byte() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let mut result = String::new();
                    for i in start..self.stack.len() {
                        result.push_str(&format!("{}", self.stack[i]));
                    }
                    self.stack.truncate(start);
                    self.push(Value::String(Arc::from(result.as_str())))?;
                }

                // -- Bitwise --
                Op::i32_and => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a & b))?; }
                Op::i32_or  => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a | b))?; }
                Op::i32_xor => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a ^ b))?; }
                Op::i32_not => { let a = self.pop().as_i32(); self.push(Value::I32(!a))?; }
                Op::i32_shl    => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a << (b & 0x1f)))?; }
                Op::i32_shr_s    => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a >> (b & 0x1f)))?; }
                Op::i32_shr_u   => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32((a >> (b & 0x1f)) as i32))?; }

                // -- Comparison --
                Op::eq => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.eq(&b)))?;
                }
                Op::ne => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(!a.eq(&b)))?;
                }
                Op::f64_lt => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a < b))?; }
                Op::f64_gt => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a > b))?; }
                Op::f64_le => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a <= b))?; }
                Op::f64_ge => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a >= b))?; }
                Op::str_lt => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.as_str() < b.as_str()))?;
                }
                Op::str_gt => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.as_str() > b.as_str()))?;
                }

                // -- Logical --
                Op::bool_not => {
                    let a = self.pop().as_bool();
                    self.push(Value::Bool(!a))?;
                }

                // -- Control flow --
                Op::br => {
                    let offset = self.read_i16();
                    let f = self.frame_mut();
                    f.ip = (f.ip as i64 + offset as i64) as usize;
                }
                Op::br_if_false => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == false {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                Op::br_if_true => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == true {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                Op::br_if_null => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if matches!(val, Value::Null) {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }

                // -- Functions --
                Op::call => {
                    let argc = self.read_byte() as usize;
                    self.call_value(argc)?;
                }
                Op::call_ref => {
                    // Direct call through a function reference — same as call
                    // but the func ref is already on the stack (no table lookup).
                    let argc = self.read_byte() as usize;
                    self.call_value(argc)?;
                }
                Op::r#return => {
                    let result = self.pop();
                    let base = self.frame().base;
                    self.close_upvalues(base);
                    self.frames.pop();
                    if self.frames.is_empty() || self.frames.len() < min_depth {
                        return Ok(result);
                    }
                    self.stack.truncate(base);
                    self.push(result)?;
                }
                Op::ref_func => {
                    let func_idx = self.read_u16() as usize;
                    let chunk = &self.chunks[func_idx];
                    let arity = chunk.arity;
                    let name = if chunk.name == "<script>" { None } else { Some(chunk.name.clone()) };

                    let uv_count = self.read_byte() as usize;
                    let mut upvalues: Vec<Arc<Mutex<Upvalue>>> = Vec::with_capacity(uv_count);
                    for _ in 0..uv_count {
                        let is_local = self.read_byte() != 0;
                        let index = self.read_byte() as usize;
                        if is_local {
                            let base = self.frame().base;
                            let uv = self.capture_upvalue(base + index);
                            upvalues.push(uv);
                        } else {
                            let uv = self.frame().upvalues[index].clone();
                            upvalues.push(uv);
                        }
                    }

                    let func = Function { name, arity, chunk_index: func_idx, upvalues };
                    let mut obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0, fields: Vec::new() };
                    // Add to function table for call_indirect
                    let table_idx = self.func_table.len();
                    obj.properties.insert("__table_idx".into(), Value::F64(table_idx as f64));
                    let func_val = Value::Object(Arc::new(Mutex::new(obj)));
                    self.func_table.push(func_val.clone());
                    self.push(func_val)?;
                }

                // -- Host functions --
                Op::call_import => {
                    let import_idx = self.read_u16() as usize;
                    let argc = self.read_byte() as usize;

                    if import_idx >= self.import_table.len() {
                        return Err(VMError::new(format!("Unresolved import index: {}", import_idx)));
                    }

                    match self.import_table[import_idx].clone() {
                        ImportTarget::Host(host_idx) => {
                            let base = self.stack.len() - argc;
                            let args: Vec<Value> = self.stack[base..].to_vec();
                            self.stack.truncate(base);

                            let placeholder: HostFn = Arc::new(|_, _| Value::Null);
                            let host_fn = std::mem::replace(&mut self.host_fns[host_idx], placeholder);
                            let result = {
                                let mut ctx = self.make_host_context();
                                host_fn(&mut ctx, &args)
                            };
                            self.host_fns[host_idx] = host_fn;

                            // JSPI: transparent async suspension
                            if let Value::Object(ref obj) = result {
                                let o = obj.lock().unwrap();
                                let is_pending = o.properties.get("__type")
                                    .map(|v| format!("{}", v) == "Promise")
                                    .unwrap_or(false)
                                    && o.properties.get("__state")
                                        .map(|v| format!("{}", v) == "pending")
                                        .unwrap_or(false);
                                if is_pending {
                                    let promise_id = o.properties.get("__id")
                                        .map(|v| v.as_f64() as u64)
                                        .unwrap_or(0);
                                    drop(o);
                                    let fiber = self.save_fiber();
                                    self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                                    return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                                }
                            }

                            self.push(result)?;
                        }
                        ImportTarget::ChunkFn { chunk_index, arity } => {
                            // Component-exported function: build Function value, push below args, call.
                            let func = crate::value::Function {
                                name: None,
                                arity,
                                chunk_index,
                                upvalues: Vec::new(),
                            };
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::Function(func);
                            let func_val = Value::Object(Arc::new(Mutex::new(obj)));
                            let args_start = self.stack.len() - argc;
                            self.stack.insert(args_start, func_val);
                            self.call_value(argc)?;
                        }
                        ImportTarget::StdlibRedirect(ref global_name) => {
                            if let Some(func_val) = self.globals.get(global_name).cloned() {
                                let args_start = self.stack.len() - argc;
                                self.stack.insert(args_start, func_val);
                                self.call_value(argc)?;
                            } else {
                                return Err(VMError::new(format!(
                                    "Stdlib redirect not found: {}", global_name
                                )));
                            }
                        }
                    }
                }

                // -- Object/Array --
                Op::struct_new => {
                    let count = self.read_u16() as usize;
                    let mut obj = Object::new();
                    let needed = count * 2;
                    let available = self.stack.len();
                    let start = if needed <= available { available - needed } else { 0 };
                    for i in 0..count {
                        let key = format!("{}", self.stack[start + i * 2]);
                        let val = self.stack[start + i * 2 + 1].clone();
                        obj.set(key, val);
                    }
                    self.stack.truncate(start);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                Op::array_new => {
                    let count = self.read_u16() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
                }

                // -- Immediates --
                Op::null => self.push(Value::Null)?,
                Op::undefined => self.push(Value::Undefined)?,
                Op::r#true => self.push(Value::Bool(true))?,
                Op::r#false => self.push(Value::Bool(false))?,
                Op::i32_const_0 => self.push(Value::I32(0))?,
                Op::i32_const_1 => self.push(Value::I32(1))?,
                Op::f64_const_0 => self.push(Value::F64(0.0))?,

                // -- Type checks --
                // ref_test: TypeOf...Is using TypeRegistry hierarchy.
                // Delegates to test_type() which handles: type_id lookup,
                // __type string, __types array (JS inheritance), __control_type.
                Op::ref_test => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.pop();
                    let result = self.test_type(&val, &target_name);
                    self.push(Value::Bool(result))?;
                }
                Op::ref_cast => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    let is_type = self.test_type(&val, &target_name);
                    if !is_type {
                        return Err(VMError::new(&format!("ref.cast failed: value is not {}", target_name)));
                    }
                    // Value stays on stack (cast is a no-op if it passes)
                }
                Op::br_on_cast => {
                    let type_name_idx = self.read_u16();
                    let offset = self.read_i16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if self.test_type(&val, &target_name) {
                        // Type matches: branch (value stays on stack)
                        let ip = self.frame().ip as i64 + offset as i64;
                        self.frame_mut().ip = ip as usize;
                    }
                    // Type doesn't match: fall through (value stays on stack)
                }
                Op::br_on_cast_fail => {
                    let type_name_idx = self.read_u16();
                    let offset = self.read_i16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if !self.test_type(&val, &target_name) {
                        let ip = self.frame().ip as i64 + offset as i64;
                        self.frame_mut().ip = ip as usize;
                    }
                }

                // -- i31ref (tagged small integers) --
                Op::i31_new => {
                    // Box i32 as i31ref. In our VM, I32 is already unboxed,
                    // so this is a no-op identity. The optimization is that
                    // the VM can use I32 directly without heap allocation.
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?; // mask to 31 bits
                }
                Op::i31_get_s => {
                    let v = self.pop().as_i32();
                    // Sign extend from 31 bits
                    let extended = if v & 0x4000_0000 != 0 {
                        v | !0x7FFF_FFFF_u32 as i32
                    } else { v };
                    self.push(Value::I32(extended))?;
                }
                Op::i31_get_u => {
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?;
                }

                Op::ref_is_null => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Null | Value::Undefined)))?; }
                Op::ref_is_string => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::String(_))))?; }
                Op::ref_is_number => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::F64(_) | Value::I32(_) | Value::I64(_))))?; }
                Op::ref_is_bool => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Bool(_))))?; }
                Op::ref_is_object => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Object(_))))?; }
                Op::ref_is_func => {
                    let v = self.pop();
                    let is_fn = matches!(&v, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Function(_)));
                    self.push(Value::Bool(is_fn))?;
                }

                // -- Conversions --
                Op::f64_from_i32 => {
                    let v = self.pop();
                    self.push(Value::F64(v.as_f64()))?;
                }
                Op::i32_from_f64 => {
                    let v = self.pop();
                    self.push(Value::I32(v.as_i32()))?;
                }

                // -- Dynamic ops (inline type dispatch, no host call) --
                Op::dyn_add => {
                    let b = self.pop();
                    let a = self.pop();
                    let result = match (&a, &b) {
                        (Value::F64(x), Value::F64(y)) => Value::F64(x + y),
                        (Value::I32(x), Value::I32(y)) => Value::I32(x.wrapping_add(*y)),
                        (Value::String(_), _) | (_, Value::String(_)) => {
                            Value::String(Arc::from(format!("{}{}", a, b).as_str()))
                        }
                        // Object with __add__ dunder → cross-language operator overloading
                        (Value::Object(obj), _) => {
                            self.try_dunder_binary(obj, &b, "__add__")
                                .unwrap_or_else(|| Value::F64(a.as_f64() + b.as_f64()))
                        }
                        _ => Value::F64(a.as_f64() + b.as_f64()),
                    };
                    self.push(result)?;
                }
                Op::dyn_eq => {
                    let b = self.pop();
                    let a = self.pop();
                    let result = match (&a, &b) {
                        // null == null, undefined == undefined, null == undefined
                        (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => true,
                        (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => true,
                        (Value::Bool(x), Value::Bool(y)) => x == y,
                        (Value::F64(x), Value::F64(y)) => if x.is_nan() || y.is_nan() { false } else { x == y },
                        (Value::I32(x), Value::I32(y)) => x == y,
                        (Value::F64(x), Value::I32(y)) => *x == *y as f64,
                        (Value::I32(x), Value::F64(y)) => *x as f64 == *y,
                        // String == Number coercion
                        (Value::String(s), Value::F64(n)) | (Value::F64(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv == *n } else { false }
                        }
                        (Value::String(s), Value::I32(n)) | (Value::I32(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv == *n as f64 } else { false }
                        }
                        (Value::String(x), Value::String(y)) => x == y,
                        (Value::Object(x), Value::Object(y)) => Arc::ptr_eq(x, y),
                        _ => false,
                    };
                    self.push(Value::Bool(result))?;
                }
                Op::dyn_ne => {
                    let b = self.pop();
                    let a = self.pop();
                    let result = match (&a, &b) {
                        (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => false,
                        (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => false,
                        (Value::Bool(x), Value::Bool(y)) => x != y,
                        (Value::F64(x), Value::F64(y)) => if x.is_nan() || y.is_nan() { true } else { x != y },
                        (Value::I32(x), Value::I32(y)) => x != y,
                        (Value::F64(x), Value::I32(y)) => *x != *y as f64,
                        (Value::I32(x), Value::F64(y)) => *x as f64 != *y,
                        (Value::String(s), Value::F64(n)) | (Value::F64(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv != *n } else { true }
                        }
                        (Value::String(s), Value::I32(n)) | (Value::I32(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv != *n as f64 } else { true }
                        }
                        (Value::String(x), Value::String(y)) => x != y,
                        (Value::Object(x), Value::Object(y)) => !Arc::ptr_eq(x, y),
                        _ => true,
                    };
                    self.push(Value::Bool(result))?;
                }
                Op::dyn_lt => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x < *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__lt__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() < b.as_f64()),
                        _ => a.as_f64() < b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_gt => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x > *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__gt__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() > b.as_f64()),
                        _ => a.as_f64() > b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_le => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x <= *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__le__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() <= b.as_f64()),
                        _ => a.as_f64() <= b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_ge => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x >= *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__ge__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() >= b.as_f64()),
                        _ => a.as_f64() >= b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_neg => {
                    let a = self.pop();
                    self.push(Value::F64(-a.as_f64()))?;
                }
                Op::dyn_not => {
                    let a = self.pop();
                    self.push(Value::Bool(!dyn_truthy(&a)))?;
                }
                Op::dyn_to_bool => {
                    let a = self.pop();
                    self.push(Value::Bool(dyn_truthy(&a)))?;
                }

                // -- Async (await) --
                Op::r#await => {
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        let is_promise = o.properties.get("__type")
                            .map(|v| format!("{}", v) == "Promise")
                            .unwrap_or(false);
                        if is_promise {
                            let state = o.properties.get("__state")
                                .map(|v| format!("{}", v))
                                .unwrap_or_default();
                            if state == "fulfilled" {
                                // Already resolved — push the value and continue
                                let resolved = o.properties.get("__value").cloned().unwrap_or(Value::Null);
                                drop(o);
                                self.push(resolved)?;
                            } else if state == "pending" {
                                // Not yet resolved — suspend the fiber
                                let promise_id = o.properties.get("__id")
                                    .map(|v| v.as_f64() as u64)
                                    .unwrap_or(0);
                                drop(o);
                                let fiber = self.save_fiber();
                                self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                                // Signal suspension via special error
                                return Err(VMError::new(format!("__await__:{}", promise_id)));
                            } else {
                                // Rejected — push the rejection value
                                let rejected = o.properties.get("__value").cloned().unwrap_or(Value::Null);
                                drop(o);
                                self.push(rejected)?;
                            }
                        } else {
                            drop(o);
                            // Not a Promise — await on non-Promise returns the value as-is
                            self.push(val)?;
                        }
                    } else {
                        // Not an object — return as-is
                        self.push(val)?;
                    }
                }

                Op::set_timer => {
                    let ms = self.pop().as_f64();
                    let callback = self.pop();
                    self.event_loop.borrow_mut().queue_timer(callback, ms);
                    self.push(Value::Null)?;
                }

                // -- Exceptions (WASM exception proposal) --
                Op::try_start => {
                    let catch_offset = self.read_u16() as i16;
                    let _finally_offset = self.read_u16(); // reserved for finally
                    let f = self.frame();
                    let catch_ip = (f.ip as i64 + catch_offset as i64) as usize;
                    self.exception_handlers.push(ExceptionHandler {
                        catch_ip,
                        _chunk_index: f.chunk_index,
                        stack_depth: self.stack.len(),
                        frame_depth: self.frames.len(),
                        tag: 0, // catch-all
                    });
                }
                Op::try_end => {
                    // Normal exit from try block — pop the handler
                    self.exception_handlers.pop();
                }
                Op::throw | Op::throw_ref => {
                    let val = self.pop();
                    // Find a matching handler by walking the handler stack.
                    // Tag 0 = catch-all (always matches).
                    // Tag N = typed catch — match if exception type matches exception_tags[N].
                    let mut matched_idx = None;
                    for i in (0..self.exception_handlers.len()).rev() {
                        let handler = &self.exception_handlers[i];
                        if handler.tag == 0 {
                            // Catch-all — always matches
                            matched_idx = Some(i);
                            break;
                        }
                        // Typed catch — check if thrown value's type matches the tag
                        let tag_idx = handler.tag as usize;
                        let tag_name = self.chunks.get(0)
                            .and_then(|c| c.exception_tags.get(tag_idx))
                            .cloned()
                            .unwrap_or_default();
                        if !tag_name.is_empty() {
                            // Check: is val an instance of tag_name?
                            let matches = self.test_type(&val, &tag_name.to_lowercase())
                                || self.exception_value_matches(&val, &tag_name);
                            if matches {
                                matched_idx = Some(i);
                                break;
                            }
                        }
                        // This handler doesn't match — keep looking
                    }

                    if let Some(idx) = matched_idx {
                        // Remove this handler and all handlers above it
                        let handler = self.exception_handlers[idx].clone();
                        self.exception_handlers.truncate(idx);
                        // Unwind: restore stack and frames
                        while self.frames.len() > handler.frame_depth {
                            let base = self.frames.last().unwrap().base;
                            self.close_upvalues(base);
                            self.frames.pop();
                        }
                        self.stack.truncate(handler.stack_depth);
                        self.push(val)?;
                        let f = self.frame_mut();
                        f.ip = handler.catch_ip;
                    } else {
                        let stack = self.capture_call_stack();
                        return Err(VMError::new(format!("{}", val)).with_stack(stack));
                    }
                }
                Op::try_table => {
                    // WASM EH Phase 4: [try_table, u8 handler_count, then for each: u8 tag, u16 offset]
                    // Tag 0 = catch-all. Tag N = typed catch for exception_tags[N].
                    // Handlers are pushed in reverse order so the most specific (first) handler
                    // is on top of the stack and checked first during throw.
                    let handler_count = self.read_byte() as usize;
                    let mut handlers = Vec::new();
                    for _ in 0..handler_count {
                        let tag = self.read_byte();
                        let offset = self.read_u16();
                        let ip = self.frame().ip + offset as usize;
                        handlers.push(ExceptionHandler {
                            catch_ip: ip,
                            stack_depth: self.stack.len(),
                            frame_depth: self.frames.len(),
                            _chunk_index: self.frame().chunk_index,
                            tag,
                        });
                    }
                    // Push in reverse so first handler is checked first (it's on top)
                    for h in handlers.into_iter().rev() {
                        self.exception_handlers.push(h);
                    }
                }

                // -- Tail call --
                Op::return_call => {
                    let argc = self.read_byte() as usize;
                    // Reuse current frame: move args to base, reset IP
                    let base = self.frame().base;
                    let args_start = self.stack.len() - argc;
                    // Copy callee + args down to base
                    let callee_idx = args_start - 1;
                    for i in 0..=argc {
                        self.stack[base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(base + 1 + argc);
                    // Pop current frame and call
                    self.frames.pop();
                    self.call_value(argc)?;
                }
                Op::return_call_indirect => {
                    let argc = self.read_byte() as usize;
                    // Stack: [..., func_table_idx, arg0, arg1, ..., argN]
                    // table_idx is BELOW the args
                    let args_start = self.stack.len() - argc;
                    let table_idx_pos = args_start - 1;
                    let table_idx = self.stack[table_idx_pos].as_i32() as usize;
                    if table_idx < self.func_table.len() {
                        let func = self.func_table[table_idx].clone();
                        // Replace table_idx with the resolved function
                        self.stack[table_idx_pos] = func;
                        let base = self.frame().base;
                        let callee_idx = table_idx_pos;
                        for i in 0..=argc {
                            self.stack[base + i] = self.stack[callee_idx + i].clone();
                        }
                        self.stack.truncate(base + 1 + argc);
                        self.frames.pop();
                        self.call_value(argc)?;
                    }
                }
                Op::return_call_ref => {
                    // Same as return_call — func ref is already on stack
                    let argc = self.read_byte() as usize;
                    let base = self.frame().base;
                    let args_start = self.stack.len() - argc;
                    let callee_idx = args_start - 1;
                    for i in 0..=argc {
                        self.stack[base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(base + 1 + argc);
                    self.frames.pop();
                    self.call_value(argc)?;
                }

                // -- Linear memory --
                Op::memory_size => {
                    let pages = (self.active_mem_len() / 65536) as i32;
                    self.push(Value::I32(pages))?;
                }
                Op::memory_grow => {
                    let pages = self.pop().as_f64() as usize;
                    let old_pages = self.active_mem_grow(pages);
                    self.push(Value::I32(old_pages as i32))?;
                }
                Op::i32_load => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I32(self.memory.load_i32(addr)?))?;
                }
                Op::i32_store => {
                    let val = self.pop().as_f64() as i32;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_i32(addr, val)?;
                }
                Op::i64_load => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I64(self.memory.load_i64(addr)?))?;
                }
                Op::i64_store => {
                    let val = self.pop().as_f64() as i64;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_i64(addr, val)?;
                }
                Op::f64_load => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::F64(self.memory.load_f64(addr)?))?;
                }
                Op::f64_store => {
                    let val = self.pop().as_f64();
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_f64(addr, val)?;
                }
                Op::i32_load8_u => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I32(self.memory.load_u8(addr)? as i32))?;
                }
                Op::i32_store8 => {
                    let val = self.pop().as_f64() as u8;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_u8(addr, val)?;
                }
                Op::f32_load => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 4 <= buf.len() {
                            f32::from_le_bytes(buf[addr..addr+4].try_into().unwrap()) as f64
                        } else { 0.0 }
                    });
                    self.push(Value::F64(val))?;
                }
                Op::f32_store => {
                    let val = self.pop().as_f64() as f32;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 4 <= buf.len() {
                            buf[addr..addr+4].copy_from_slice(&val.to_le_bytes());
                        }
                    });
                }
                Op::i32_load8_s => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I32(self.memory.load_u8(addr)? as i8 as i32))?;
                }
                Op::i32_load16_s => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { i16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i32 } else { 0 }
                    });
                    self.push(Value::I32(val))?;
                }
                Op::i32_load16_u => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { u16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i32 } else { 0 }
                    });
                    self.push(Value::I32(val))?;
                }
                Op::i32_store16 => {
                    let val = self.pop().as_i32() as i16;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 2 <= buf.len() { buf[addr..addr+2].copy_from_slice(&val.to_le_bytes()); }
                    });
                }
                Op::i64_load8_s => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_u8(addr)? as i8 as i64))?;
                }
                Op::i64_load8_u => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_u8(addr)? as i64))?;
                }
                Op::i64_load16_s => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { i16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i64 } else { 0 }
                    });
                    self.push(Value::I64(val))?;
                }
                Op::i64_load16_u => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { u16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i64 } else { 0 }
                    });
                    self.push(Value::I64(val))?;
                }
                Op::i64_load32_s => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_i32(addr)? as i64))?;
                }
                Op::i64_load32_u => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_i32(addr)? as u32 as i64))?;
                }
                Op::i64_store8 => {
                    let val = self.pop().as_i64() as u8;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.store_u8(addr, val)?;
                }
                Op::i64_store16 => {
                    let val = self.pop().as_i64() as i16;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 2 <= buf.len() { buf[addr..addr+2].copy_from_slice(&val.to_le_bytes()); }
                    });
                }
                Op::i64_store32 => {
                    let val = self.pop().as_i64() as i32;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.store_i32(addr, val)?;
                }

                // -- Conversions --
                Op::i32_wrap_i64 => { let a = self.pop().as_i64(); self.push(Value::I32(a as i32))?; }
                Op::i64_extend_i32_s => { let a = self.pop().as_i32(); self.push(Value::I64(a as i64))?; }
                Op::i64_extend_i32_u => { let a = self.pop().as_i32() as u32; self.push(Value::I64(a as i64))?; }
                Op::i64_trunc_f64_s => { let a = self.pop().as_f64(); self.push(Value::I64(a as i64))?; }
                Op::i64_trunc_f64_u => { let a = self.pop().as_f64(); self.push(Value::I64(a as u64 as i64))?; }
                Op::f64_promote_f32 => { let a = self.pop().as_f64(); self.push(Value::F64(a))?; }
                Op::f32_demote_f64 => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32) as f64))?; }
                Op::i32_reinterpret_f32 => { let a = self.pop().as_f64() as f32; self.push(Value::I32(a.to_bits() as i32))?; }
                Op::i64_reinterpret_f64 => { let a = self.pop().as_f64(); self.push(Value::I64(a.to_bits() as i64))?; }
                Op::f32_reinterpret_i32 => { let a = self.pop().as_i32(); self.push(Value::F64(f32::from_bits(a as u32) as f64))?; }
                Op::f64_reinterpret_i64 => { let a = self.pop().as_i64(); self.push(Value::F64(f64::from_bits(a as u64)))?; }

                // -- Sign extension --
                Op::i32_extend8_s => { let a = self.pop().as_i32() as i8; self.push(Value::I32(a as i32))?; }
                Op::i32_extend16_s => { let a = self.pop().as_i32() as i16; self.push(Value::I32(a as i32))?; }
                Op::i64_extend8_s => { let a = self.pop().as_i64() as i8; self.push(Value::I64(a as i64))?; }
                Op::i64_extend16_s => { let a = self.pop().as_i64() as i16; self.push(Value::I64(a as i64))?; }
                Op::i64_extend32_s => { let a = self.pop().as_i64() as i32; self.push(Value::I64(a as i64))?; }

                // -- Multi-value --
                Op::pack => {
                    let count = self.read_byte() as usize;
                    let start = self.stack.len() - count;
                    let values: Vec<Value> = self.stack.drain(start..).collect();
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(values)))))?;
                }
                Op::unpack => {
                    let arr = self.pop();
                    if let Value::Object(obj) = arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = o.kind {
                            let elems = elems.clone();
                            drop(o);
                            for elem in elems {
                                self.push(elem)?;
                            }
                        }
                    }
                }

                // -- Block/loop structured control --
                Op::block => {
                    let end_offset = self.read_u16() as usize;
                    let ip = self.frame().ip;
                    self.label_stack.push(LabelEntry { target: ip + end_offset, is_loop: false });
                }
                Op::r#loop => {
                    let _body_size = self.read_u16();
                    let ip = self.frame().ip;
                    // Loop target is the start (current position, after reading the operand)
                    self.label_stack.push(LabelEntry { target: ip, is_loop: true });
                }
                Op::end => {
                    self.label_stack.pop();
                }
                Op::br_label => {
                    let depth = self.read_byte() as usize;
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                        let target = entry.target;
                        let _ci = self.frame().chunk_index;
                        self.frames.last_mut().unwrap().ip = target;
                        // If branching out of a block (not loop), pop labels
                        if !entry.is_loop {
                            let len = self.label_stack.len();
                            self.label_stack.truncate(len - depth - 1);
                        }
                    }
                }
                Op::br_if_label => {
                    let depth = self.read_byte() as usize;
                    let cond = self.pop();
                    if dyn_truthy(&cond) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                            let target = entry.target;
                            self.frames.last_mut().unwrap().ip = target;
                            if !entry.is_loop {
                                let len = self.label_stack.len();
                                self.label_stack.truncate(len - depth - 1);
                            }
                        }
                    }
                }
                Op::br_table => {
                    let count = self.read_byte() as usize;
                    let default_depth = self.read_byte() as usize;
                    let mut labels = Vec::with_capacity(count);
                    for _ in 0..count { labels.push(self.read_byte() as usize); }
                    let idx = self.pop().as_f64() as usize;
                    let depth = if idx < count { labels[idx] } else { default_depth };
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                        let target = entry.target;
                        self.frames.last_mut().unwrap().ip = target;
                    }
                }

                // -- call_indirect --
                Op::call_indirect => {
                    let argc = self.read_byte() as usize;
                    let table_idx_pos = self.stack.len() - 1 - argc;
                    let raw_idx = self.stack[table_idx_pos].as_f64();
                    if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= self.func_table.len() as f64 {
                        return Err(VMError::new(format!("trap: call_indirect: invalid table index {}", raw_idx)));
                    }
                    let table_idx = raw_idx as usize;
                    if table_idx < self.func_table.len() {
                        self.stack[table_idx_pos] = self.func_table[table_idx].clone();
                        self.call_value(argc)?;
                    } else {
                        return Err(VMError::new(format!("call_indirect: table index {} out of bounds", table_idx)));
                    }
                }

                // -- Component Model --
                Op::canon_lift => {
                    let type_idx = self.read_u16() as usize;
                    // Lift: convert core value to component interface type.
                    // For now: if value is an object, stamp its type_id.
                    // In full CM, this would validate/convert the value shape.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let mut o = obj.lock().unwrap();
                        if o.type_id == 0 && type_idx < self.type_registry.types.len() {
                            o.type_id = type_idx;
                        }
                    }
                    self.push(val)?;
                }
                Op::canon_lower => {
                    let type_idx = self.read_u16() as usize;
                    // Lower: convert component interface type to core value.
                    // For now: validate type_id matches, strip interface metadata.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if type_idx < self.type_registry.types.len() && o.type_id != type_idx {
                            // Type mismatch — could trap, for now allow
                        }
                    }
                    self.push(val)?;
                }
                Op::type_import => {
                    let import_idx = self.read_u16() as usize;
                    // Resolve type import: look up in type_imports table and register.
                    // The type should already be registered during linking.
                    // Push the type_id onto the stack for use by constructors.
                    if import_idx < self.chunks[0].type_imports.len() {
                        let (_iface, type_name) = &self.chunks[0].type_imports[import_idx];
                        if let Some(tid) = self.type_registry.get_id(type_name) {
                            self.push(Value::I32(tid as i32))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                Op::type_export => {
                    let type_id = self.read_u16() as usize;
                    // Mark type as exported. This is a compile-time declaration;
                    // at runtime, ensure the type is visible to the linker.
                    // No stack effect.
                    let _ = type_id;
                }
                Op::shared_new => {
                    // Create a new shared object from the type_id on stack.
                    let type_id_val = self.pop();
                    let type_id = type_id_val.as_f64() as usize;
                    let mut obj = Object::new();
                    obj.type_id = type_id;
                    if let Some(td) = self.type_registry.get(type_id) {
                        obj.fields = vec![Value::Null; td.field_count()];
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }

                // -- Shared-Everything Threads (shared GC objects) --
                Op::shared_struct_get => {
                    let field_idx = self.read_u16() as usize;
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let o = obj.lock().unwrap();
                        // Atomic read of indexed field
                        let val = if field_idx < o.fields.len() {
                            o.fields[field_idx].clone()
                        } else {
                            Value::Null
                        };
                        self.push(val)?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                Op::shared_struct_set => {
                    let field_idx = self.read_u16() as usize;
                    let value = self.pop();
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let mut o = obj.lock().unwrap();
                        // Atomic write of indexed field
                        if field_idx < o.fields.len() {
                            o.fields[field_idx] = value;
                        } else {
                            // Grow fields if needed
                            while o.fields.len() <= field_idx {
                                o.fields.push(Value::Null);
                            }
                            o.fields[field_idx] = value;
                        }
                    }
                }
                Op::shared_array_get => {
                    let idx_val = self.pop();
                    let arr_val = self.pop();
                    let idx = idx_val.as_f64() as usize;
                    if let Value::Object(ref obj) = arr_val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = o.kind {
                            self.push(elems.get(idx).cloned().unwrap_or(Value::Null))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                Op::shared_array_set => {
                    let value = self.pop();
                    let idx_val = self.pop();
                    let arr_val = self.pop();
                    let idx = idx_val.as_f64() as usize;
                    if let Value::Object(ref obj) = arr_val {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut elems) = o.kind {
                            if idx < elems.len() {
                                elems[idx] = value;
                            }
                        }
                    }
                }
                Op::shared_struct_cas => {
                    let field_idx = self.read_u16() as usize;
                    let new_val = self.pop();
                    let expected = self.pop();
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let mut o = obj.lock().unwrap();
                        if field_idx < o.fields.len() {
                            let old = o.fields[field_idx].clone();
                            // Compare (using string repr for simplicity)
                            if format!("{}", old) == format!("{}", expected) {
                                o.fields[field_idx] = new_val;
                            }
                            self.push(old)?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }

                // -- Weak References & Finalizers --
                Op::ref_make_weak => {
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        self.push(Value::WeakRef(Arc::downgrade(obj)))?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                Op::ref_deref_weak => {
                    let val = self.pop();
                    if let Value::WeakRef(ref weak) = val {
                        if let Some(strong) = weak.upgrade() {
                            self.push(Value::Object(strong))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        // If it's already a strong ref, just pass through
                        self.push(val)?;
                    }
                }
                Op::ref_is_alive => {
                    let val = self.pop();
                    let alive = match &val {
                        Value::WeakRef(weak) => weak.upgrade().is_some(),
                        Value::Object(_) => true,
                        _ => false,
                    };
                    self.push(Value::Bool(alive))?;
                }
                Op::ref_register_finalizer => {
                    let callback = self.pop();
                    let target = self.pop();
                    if let Value::Object(ref obj) = target {
                        self.finalizers.push(FinalizerEntry {
                            target: Arc::downgrade(obj),
                            callback,
                        });
                    }
                }

                // -- Multi-Memory --
                Op::memory_select => {
                    let mem_idx = self.read_byte() as usize;
                    self.active_memory = mem_idx;
                }
                Op::memory_init => {
                    let pages = self.pop().as_f64() as usize;
                    let mem_idx = self.extra_memories.len() + 1; // 0 is default memory
                    self.extra_memories.push(vec![0u8; pages * 65536]);
                    self.push(Value::I32(mem_idx as i32))?;
                }
                Op::memory_copy_cross => {
                    let len = self.pop().as_f64() as usize;
                    let src_addr = self.pop().as_f64() as usize;
                    let src_mem = self.pop().as_f64() as usize;
                    let dst_addr = self.pop().as_f64() as usize;
                    let dst_mem = self.pop().as_f64() as usize;
                    // Copy bytes between memories
                    let src_data: Vec<u8> = if src_mem == 0 {
                        let mut buf = vec![0u8; len];
                        self.memory.read_bytes(src_addr, &mut buf);
                        buf
                    } else {
                        let src = self.extra_mem(src_mem);
                        if src_addr + len <= src.len() {
                            src[src_addr..src_addr + len].to_vec()
                        } else {
                            vec![0u8; len]
                        }
                    };
                    if dst_mem == 0 {
                        self.memory.write_bytes(dst_addr, &src_data);
                    } else {
                        let dst = self.extra_mem_mut(dst_mem);
                        if dst_addr + len <= dst.len() {
                            dst[dst_addr..dst_addr + len].copy_from_slice(&src_data);
                        }
                    }
                }

                // -- JS String Builtins (wasm:js-string proposal) --
                Op::str_length => {
                    let s = self.pop();
                    let len = match &s {
                        Value::String(s) => s.chars().count() as i32,
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Array(a) = &o.kind { a.len() as i32 }
                            else if let Some(Value::F64(n)) = o.properties.get("length") { *n as i32 }
                            else { 0 }
                        }
                        _ => 0,
                    };
                    self.push(Value::I32(len))?;
                }
                Op::str_char_code_at => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let code = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| c as i32).unwrap_or(-1)
                    } else { -1 };
                    self.push(Value::I32(code))?;
                }
                Op::str_from_char_code => {
                    let code = self.pop().as_i32() as u32;
                    let ch = char::from_u32(code).unwrap_or('\0');
                    self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                }
                Op::str_char_at => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let ch = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| Arc::from(c.to_string().as_str()))
                            .unwrap_or(Arc::from(""))
                    } else { Arc::from("") };
                    self.push(Value::String(ch))?;
                }
                Op::str_substring | Op::str_slice => {
                    let end = self.pop().as_i32() as usize;
                    let start = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let result = if let Value::String(s) = &s {
                        let chars: Vec<char> = s.chars().collect();
                        let end = end.min(chars.len());
                        let start = start.min(end);
                        let sub: String = chars[start..end].iter().collect();
                        Arc::from(sub.as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(result))?;
                }
                Op::str_index_of => {
                    let needle = self.pop();
                    let haystack = self.pop();
                    let pos = match (&haystack, &needle) {
                        (Value::String(h), Value::String(n)) => {
                            h.find(n.as_ref()).map(|p| p as i32).unwrap_or(-1)
                        }
                        _ => -1,
                    };
                    self.push(Value::I32(pos))?;
                }
                Op::str_last_index_of => {
                    let needle = self.pop();
                    let haystack = self.pop();
                    let pos = match (&haystack, &needle) {
                        (Value::String(h), Value::String(n)) => {
                            h.rfind(n.as_ref()).map(|p| p as i32).unwrap_or(-1)
                        }
                        _ => -1,
                    };
                    self.push(Value::I32(pos))?;
                }
                Op::str_equals => {
                    let b = self.pop(); let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::String(a), Value::String(b)) => a == b,
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
                }
                Op::str_compare => {
                    let b = self.pop(); let a = self.pop();
                    let cmp = match (&a, &b) {
                        (Value::String(a), Value::String(b)) => {
                            match a.cmp(b) {
                                std::cmp::Ordering::Less => -1,
                                std::cmp::Ordering::Equal => 0,
                                std::cmp::Ordering::Greater => 1,
                            }
                        }
                        _ => 0,
                    };
                    self.push(Value::I32(cmp))?;
                }
                Op::str_to_upper => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.to_uppercase().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_to_lower => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.to_lowercase().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim_start => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim_start()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim_end => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim_end()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_starts_with => {
                    let prefix = self.pop(); let s = self.pop();
                    let r = match (&s, &prefix) {
                        (Value::String(s), Value::String(p)) => s.starts_with(p.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::str_ends_with => {
                    let suffix = self.pop(); let s = self.pop();
                    let r = match (&s, &suffix) {
                        (Value::String(s), Value::String(p)) => s.ends_with(p.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::str_contains => {
                    let needle = self.pop(); let s = self.pop();
                    let r = match (&s, &needle) {
                        (Value::String(s), Value::String(n)) => s.contains(n.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::str_replace => {
                    let new = self.pop(); let old = self.pop(); let s = self.pop();
                    let r = match (&s, &old, &new) {
                        (Value::String(s), Value::String(o), Value::String(n)) => {
                            Arc::from(s.replace(o.as_ref(), n.as_ref()).as_str())
                        }
                        _ => Arc::from(""),
                    };
                    self.push(Value::String(r))?;
                }
                Op::str_split => {
                    let delim = self.pop(); let s = self.pop();
                    let parts: Vec<Value> = match (&s, &delim) {
                        (Value::String(s), Value::String(d)) => {
                            s.split(d.as_ref()).map(|p| Value::String(Arc::from(p))).collect()
                        }
                        _ => vec![],
                    };
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(parts)))))?;
                }
                Op::str_repeat => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.repeat(count).as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_pad_start => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Arc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Arc::from(format!("{}{}", fill_str, s).as_str())
                        }
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_pad_end => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Arc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Arc::from(format!("{}{}", s, fill_str).as_str())
                        }
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_reverse => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.chars().rev().collect::<String>().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                // Unicode code points (beyond BMP — emoji, CJK)
                Op::str_from_code_point => {
                    let cp = self.pop().as_i32() as u32;
                    let ch = char::from_u32(cp).unwrap_or('\u{FFFD}');
                    self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                }
                Op::str_code_point_at => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let cp = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| c as i32).unwrap_or(-1)
                    } else { -1 };
                    self.push(Value::I32(cp))?;
                }
                // Bulk char code operations
                Op::str_into_char_codes => {
                    let s = self.pop();
                    let codes: Vec<Value> = if let Value::String(s) = &s {
                        s.chars().map(|c| Value::I32(c as i32)).collect()
                    } else { vec![] };
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(codes)))))?;
                }
                Op::str_from_char_codes => {
                    let arr = self.pop();
                    let s = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.iter().filter_map(|v| char::from_u32(v.as_i32() as u32)).collect::<String>()
                        } else { String::new() }
                    } else { String::new() };
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                // Type discrimination opcodes
                Op::ref_typeof => {
                    let v = self.pop();
                    let tag = match &v {
                        Value::Undefined => "undefined",
                        Value::Null => "object",  // JS spec: typeof null === "object"
                        Value::Bool(_) => "boolean",
                        Value::I32(_) | Value::I64(_) | Value::F64(_) => "number",
                        Value::String(_) => "string",
                        Value::V128(_) => "v128",
                        Value::WeakRef(_) => "weakref",
                        Value::Object(o) => {
                            let ob = o.lock().unwrap();
                            match &ob.kind {
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                                ObjectKind::Array(_) => "array",
                                _ => "object",
                            }
                        }
                    };
                    self.push(Value::String(Arc::from(tag)))?;
                }
                Op::ref_is_array => {
                    let v = self.pop();
                    let is_arr = matches!(&v, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Array(_)));
                    self.push(Value::Bool(is_arr))?;
                }

                // -- Array builtins --
                Op::array_length => {
                    let arr = self.pop();
                    let len = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind { a.len() as i32 } else { 0 }
                    } else if let Value::String(s) = &arr {
                        s.chars().count() as i32
                    } else { 0 };
                    self.push(Value::I32(len))?;
                }
                Op::array_push => {
                    let val = self.pop(); let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.push(val); }
                    }
                    self.push(arr)?;
                }
                Op::array_pop => {
                    let arr = self.pop();
                    let val = if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.pop().unwrap_or(Value::Null) }
                        else { Value::Null }
                    } else { Value::Null };
                    self.push(val)?;
                }
                Op::array_slice => {
                    let end = self.pop().as_i32(); let start = self.pop().as_i32();
                    let arr = self.pop();
                    let result = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            let len = a.len() as i32;
                            let s = if start < 0 { (len + start).max(0) as usize } else { start.min(len) as usize };
                            let e = if end < 0 { (len + end).max(0) as usize } else { end.min(len) as usize };
                            let sliced: Vec<Value> = a[s..e.max(s)].to_vec();
                            Value::Object(Arc::new(Mutex::new(Object::new_array(sliced))))
                        } else { Value::Null }
                    } else { Value::Null };
                    self.push(result)?;
                }
                Op::array_join => {
                    let delim = self.pop(); let arr = self.pop();
                    let r = if let (Value::Object(obj), Value::String(d)) = (&arr, &delim) {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            let parts: Vec<String> = a.iter().map(|v| format!("{}", v)).collect();
                            Arc::from(parts.join(d.as_ref()).as_str())
                        } else { Arc::from("") }
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::array_reverse => {
                    let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.reverse(); }
                    }
                    self.push(arr)?;
                }
                Op::array_contains => {
                    // Compare compiles: left(needle) then right(haystack)
                    // Stack: [needle, haystack]. Pop: haystack (TOS), needle.
                    let haystack = self.pop();
                    let needle = self.pop();
                    let found = match (&haystack, &needle) {
                        // String containment: "lo" in "hello"
                        (Value::String(h), Value::String(n)) => h.contains(n.as_ref()),
                        // Array containment: 2 in [1,2,3]
                        (Value::Object(obj), _) => {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Array(a) = &o.kind {
                                a.iter().any(|v| v.eq(&needle))
                            } else {
                                // Dict/object containment: check if key exists
                                let key = format!("{}", needle);
                                o.properties.contains_key(&key)
                            }
                        }
                        _ => false,
                    };
                    self.push(Value::Bool(found))?;
                }
                Op::array_index_of => {
                    let needle = self.pop(); let arr = self.pop();
                    let idx = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.iter().position(|v| v.eq(&needle)).map(|p| p as i32).unwrap_or(-1)
                        } else { -1 }
                    } else { -1 };
                    self.push(Value::I32(idx))?;
                }

                // WASM GC array ops
                Op::array_new_default => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
                }
                Op::array_fill => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let start = self.pop().as_i32().max(0) as usize;
                    let val = self.pop();
                    let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            let end = (start + count).min(a.len());
                            for i in start..end { a[i] = val.clone(); }
                        }
                    }
                }
                Op::array_copy => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let src_off = self.pop().as_i32().max(0) as usize;
                    let src = self.pop();
                    let dst_off = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop();
                    // Read source slice
                    let src_vals: Vec<Value> = if let Value::Object(obj) = &src {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            let end = (src_off + len).min(a.len());
                            a[src_off.min(a.len())..end].to_vec()
                        } else { vec![] }
                    } else { vec![] };
                    // Write to destination
                    if let Value::Object(obj) = &dst {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            for (i, v) in src_vals.into_iter().enumerate() {
                                let idx = dst_off + i;
                                if idx < a.len() { a[idx] = v; }
                            }
                        }
                    }
                }
                Op::array_concat => {
                    let b = self.pop();
                    let a = self.pop();
                    let mut result = Vec::new();
                    if let Value::Object(obj) = &a {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(arr) = &o.kind { result.extend(arr.iter().cloned()); }
                    }
                    if let Value::Object(obj) = &b {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(arr) = &o.kind { result.extend(arr.iter().cloned()); }
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(result)))))?;
                }
                Op::array_shift => {
                    let arr = self.pop();
                    let val = if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            if a.is_empty() { Value::Null } else { a.remove(0) }
                        } else { Value::Null }
                    } else { Value::Null };
                    self.push(val)?;
                }

                // -- Stack Switching (wasm stack-switching proposal) --
                Op::cont_new => {
                    // Create a continuation from a function reference.
                    // The continuation wraps a function + saved state.
                    let func_val = self.pop();
                    let mut obj = Object::new_typed(0);
                    obj.properties.insert("__cont_func".into(), func_val);
                    obj.properties.insert("__cont_state".into(), Value::String(Arc::from("ready")));
                    obj.properties.insert("__cont_value".into(), Value::Null);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                Op::suspend => {
                    let _tag = self.read_u16();
                    // Yield a value from the current continuation.
                    // The yielded value stays on the stack for the caller.
                    // This is like a return but the continuation can be resumed.
                    let val = self.pop();
                    return Ok(val);
                }
                Op::resume => {
                    let _tag = self.read_u16();
                    // Resume a continuation, passing a value to it.
                    // [continuation, value] → [result]
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        let func_val = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_func").cloned().unwrap_or(Value::Null)
                        };
                        {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                            o.properties.insert("__cont_value".into(), val.clone());
                        }
                        // Call the continuation's function with the resume value
                        self.push(func_val)?;
                        self.push(val)?;
                        self.call_value(1)?;
                    } else {
                        self.push(val)?;
                    }
                }
                Op::switch => {
                    let _tag = self.read_u16();
                    // Symmetric switch: suspend current continuation, resume target
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        let mut o = obj.lock().unwrap();
                        o.properties.insert("__cont_value".into(), val.clone());
                        o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                    }
                    self.push(val)?;
                }

                // -- wasi-threads: real OS thread spawning --
                Op::thread_spawn => {
                    // [func_ref] → [task_object]
                    // Per wasi-threads: spawn a real OS thread.
                    // Value is now Arc-based (Send+Sync), so chunks and host_fns
                    // can be shared directly — no serialization needed.
                    let func_val = self.pop();

                    let chunk_idx = match &func_val {
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            match &o.kind {
                                ObjectKind::Function(f) => Some(f.chunk_index),
                                _ => None,
                            }
                        }
                        _ => None,
                    };

                    if let Some(ci) = chunk_idx {
                        let tid = self.next_thread_id;
                        self.next_thread_id += 1;

                        // Create task object FIRST so child can write result to it
                        let mut obj = Object::new();
                        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
                        obj.properties.insert("__thread_id".into(), Value::I32(tid));
                        obj.properties.insert("iscompleted".into(), Value::Bool(false));
                        obj.properties.insert("isalive".into(), Value::Bool(true));
                        obj.properties.insert("result".into(), Value::Null);
                        obj.properties.insert("status".into(), Value::String(Arc::from("Running")));
                        let task_obj = Arc::new(Mutex::new(obj));
                        let task_for_child = task_obj.clone();

                        // Share directly — Value is Send+Sync now
                        let child_chunks = self.chunks.clone();
                        let child_memory = self.memory.clone();
                        let child_host_fns = self.host_fns.clone();
                        let child_host_registry = self.host_registry.clone();
                        let child_import_table = self.import_table.clone();

                        let handle = std::thread::spawn(move || {
                            let mut child_vm = VM::new();
                            child_vm.chunks = child_chunks;
                            child_vm.memory = child_memory;
                            child_vm.host_fns = child_host_fns;
                            child_vm.host_registry = child_host_registry;
                            child_vm.import_table = child_import_table;

                            // Set up call frame
                            let _arity = child_vm.chunks.get(ci).map(|c| c.arity).unwrap_or(0);
                            child_vm.frames.push(CallFrame {
                                chunk_index: ci,
                                ip: 0,
                                base: 0,
                                upvalues: Vec::new(),
                            });
                            let local_count = child_vm.chunks.get(ci)
                                .map(|c| c.local_count as usize)
                                .unwrap_or(1)
                                .max(64);
                            for _ in 0..local_count {
                                child_vm.stack.push(Value::Null);
                            }

                            let result = match child_vm.execute() {
                                Ok(val) => {
                                    // Store return value in the shared task object
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("result".into(), val.clone());
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(0));
                                    t.properties.insert("status".into(), Value::String(Arc::from("RanToCompletion")));
                                    vec![0u8]
                                }
                                Err(e) => {
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(-1));
                                    t.properties.insert("status".into(), Value::String(Arc::from("Faulted")));
                                    eprintln!("[thread {}] error: {}", tid, e.message);
                                    vec![1u8]
                                }
                            };
                            result
                        });

                        self.thread_handles.insert(tid, handle);
                        self.push(Value::Object(task_obj))?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                Op::thread_join => {
                    // [task_object] → [status: i32]
                    // Wait for a thread to complete. Accepts either a task object
                    // (with __thread_id) or a raw i32 thread ID.
                    let task_val = self.pop();
                    let tid = match &task_val {
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            o.properties.get("__thread_id").map(|v| v.as_f64() as i32).unwrap_or(-1)
                        }
                        Value::I32(n) => *n,
                        _ => task_val.as_f64() as i32,
                    };

                    if let Some(handle) = self.thread_handles.remove(&tid) {
                        let success = match handle.join() {
                            Ok(result) => result.first().copied().unwrap_or(1) == 0,
                            Err(_) => false,
                        };
                        // Update the task object properties
                        if let Value::Object(obj) = &task_val {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("iscompleted".into(), Value::Bool(true));
                            o.properties.insert("isalive".into(), Value::Bool(false));
                            o.properties.insert("hasexited".into(), Value::Bool(true));
                            o.properties.insert("exitcode".into(), Value::I32(if success { 0 } else { -1 }));
                            o.properties.insert("status".into(), Value::String(Arc::from(
                                if success { "RanToCompletion" } else { "Faulted" }
                            )));
                        }
                        self.push(Value::I32(if success { 0 } else { -1 }))?;
                    } else {
                        self.push(Value::I32(-1))?;
                    }
                }

                // -- Extended Const Expressions --
                Op::global_init => {
                    let idx = self.read_u16() as usize;
                    // Evaluate the const expr for global at index idx and store result.
                    // This is a runtime opcode for cases where init can't be done at load time.
                    if idx < self.chunks[0].global_inits.len() {
                        let gi = self.chunks[0].global_inits[idx].clone();
                        let val = self.eval_const_expr(&gi.init);
                        self.globals.insert(gi.name.clone(), val);
                    }
                }

                // -- Typed Continuations --
                Op::cont_new_typed => {
                    let tag_idx = self.read_u16() as usize;
                    let func_val = self.pop();
                    let mut obj = Object::new_typed(0);
                    obj.properties.insert("__cont_func".into(), func_val);
                    obj.properties.insert("__cont_state".into(), Value::String(Arc::from("ready")));
                    obj.properties.insert("__cont_value".into(), Value::Null);
                    // Store tag info for type checking on suspend/resume
                    obj.properties.insert("__cont_tag".into(), Value::I32(tag_idx as i32));
                    if tag_idx < self.chunks[0].continuation_tags.len() {
                        let tag = &self.chunks[0].continuation_tags[tag_idx];
                        obj.properties.insert("__cont_yield_type".into(), Value::String(Arc::from(tag.yield_type.as_str())));
                        obj.properties.insert("__cont_resume_type".into(), Value::String(Arc::from(tag.resume_type.as_str())));
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                Op::suspend_typed => {
                    let _tag_idx = self.read_u16();
                    // Typed suspend: yield a value, with type validation.
                    let val = self.pop();
                    // In a full implementation, we'd validate val matches the yield_type
                    // of the current continuation's tag. For now, just yield.
                    return Ok(val);
                }
                Op::resume_typed => {
                    let _tag_idx = self.read_u16();
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        // Validate resume value type matches continuation's resume_type
                        let expected_type = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_resume_type")
                                .map(|v| format!("{}", v))
                                .unwrap_or_default()
                        };
                        if !expected_type.is_empty() && expected_type != "any" {
                            let actual_matches = self.test_type(&val, &expected_type.to_lowercase());
                            if !actual_matches {
                                // Type mismatch on resume — for now, proceed anyway
                                // A strict implementation would trap here
                            }
                        }
                        let func_val = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_func").cloned().unwrap_or(Value::Null)
                        };
                        {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                            o.properties.insert("__cont_value".into(), val.clone());
                        }
                        self.push(func_val)?;
                        self.push(val)?;
                        self.call_value(1)?;
                    } else {
                        self.push(val)?;
                    }
                }

                // -- String References (zero-copy) --
                Op::string_as_ref => {
                    // String → StringRef: since our strings are already Arc<str>,
                    // this is effectively a no-op — we keep the same Arc.
                    // The semantic difference: stringref signals cross-component sharing intent.
                    let val = self.pop();
                    // Just pass through — Arc<str> is already shared
                    self.push(val)?;
                }
                Op::string_from_ref => {
                    // StringRef → String: dereference (zero-copy with Arc).
                    let val = self.pop();
                    self.push(val)?;
                }
                Op::string_ref_eq => {
                    // Pointer identity comparison for string refs.
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::String(sa), Value::String(sb)) => Arc::ptr_eq(sa, sb),
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
                }

                // -- SIMD (128-bit vectors) --
                Op::v128_load => {
                    let addr = self.pop().as_i32() as usize;
                    let mut bytes = [0u8; 16];
                    self.memory.read_bytes(addr, &mut bytes);
                    self.push(Value::V128(bytes))?;
                }
                Op::v128_store => {
                    let val = self.pop();
                    let addr = self.pop().as_i32() as usize;
                    if let Value::V128(bytes) = val {
                        self.memory.write_bytes(addr, &bytes);
                    }
                }
                Op::v128_const => {
                    let mut bytes = [0u8; 16];
                    for i in 0..16 { bytes[i] = self.read_byte(); }
                    self.push(Value::V128(bytes))?;
                }

                // i32x4 ops
                Op::i32x4_splat => {
                    let v = self.pop().as_i32();
                    let mut bytes = [0u8; 16];
                    for i in 0..4 { bytes[i*4..i*4+4].copy_from_slice(&v.to_le_bytes()); }
                    self.push(Value::V128(bytes))?;
                }
                Op::i32x4_add => { self.simd_i32x4_binop(|a, b| a.wrapping_add(b))?; }
                Op::i32x4_sub => { self.simd_i32x4_binop(|a, b| a.wrapping_sub(b))?; }
                Op::i32x4_mul => { self.simd_i32x4_binop(|a, b| a.wrapping_mul(b))?; }
                Op::i32x4_eq =>  { self.simd_i32x4_binop(|a, b| if a == b { -1 } else { 0 })?; }
                Op::i32x4_gt_s => { self.simd_i32x4_binop(|a, b| if a > b { -1 } else { 0 })?; }
                Op::i32x4_lt_s => { self.simd_i32x4_binop(|a, b| if a < b { -1 } else { 0 })?; }
                Op::i32x4_shl => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v << shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::i32x4_shr_s => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v >> shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::i32x4_shr_u => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = u32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v >> shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::i32x4_extract_lane => {
                    let lane = self.read_byte() as usize & 3;
                    if let Value::V128(a) = self.pop() {
                        let v = i32::from_le_bytes([a[lane*4], a[lane*4+1], a[lane*4+2], a[lane*4+3]]);
                        self.push(Value::I32(v))?;
                    } else { self.push(Value::I32(0))?; }
                }
                Op::i32x4_replace_lane => {
                    let lane = self.read_byte() as usize & 3;
                    let val = self.pop().as_i32();
                    if let Value::V128(mut a) = self.pop() {
                        a[lane*4..lane*4+4].copy_from_slice(&val.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // f64x2 ops
                Op::f64x2_splat => {
                    let v = self.pop().as_f64();
                    let mut bytes = [0u8; 16];
                    bytes[0..8].copy_from_slice(&v.to_le_bytes());
                    bytes[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(bytes))?;
                }
                Op::f64x2_add => { self.simd_f64x2_binop(|a, b| a + b)?; }
                Op::f64x2_sub => { self.simd_f64x2_binop(|a, b| a - b)?; }
                Op::f64x2_mul => { self.simd_f64x2_binop(|a, b| a * b)?; }
                Op::f64x2_div => { self.simd_f64x2_binop(|a, b| a / b)?; }
                Op::f64x2_min => { self.simd_f64x2_binop(|a, b| a.min(b))?; }
                Op::f64x2_max => { self.simd_f64x2_binop(|a, b| a.max(b))?; }
                Op::f64x2_eq => { self.simd_f64x2_cmp(|a, b| a == b)?; }
                Op::f64x2_lt => { self.simd_f64x2_cmp(|a, b| a < b)?; }
                Op::f64x2_le => { self.simd_f64x2_cmp(|a, b| a <= b)?; }
                Op::f64x2_sqrt => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&v.sqrt().to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f64x2_abs => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&v.abs().to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f64x2_neg => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&(-v).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f64x2_extract_lane => {
                    let lane = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        let v = f64::from_le_bytes(a[lane*8..lane*8+8].try_into().unwrap());
                        self.push(Value::F64(v))?;
                    } else { self.push(Value::F64(0.0))?; }
                }
                Op::f64x2_replace_lane => {
                    let lane = self.read_byte() as usize & 1;
                    let val = self.pop().as_f64();
                    if let Value::V128(mut a) = self.pop() {
                        a[lane*8..lane*8+8].copy_from_slice(&val.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // f32x4, i8x16, i16x8 — same patterns, delegate to helpers
                Op::f32x4_splat => { let v = self.pop().as_f64() as f32; let b = v.to_le_bytes(); let mut out = [0u8;16]; for i in 0..4 { out[i*4..i*4+4].copy_from_slice(&b); } self.push(Value::V128(out))?; }
                Op::f32x4_add => { self.simd_f32x4_binop(|a,b| a+b)?; }
                Op::f32x4_sub => { self.simd_f32x4_binop(|a,b| a-b)?; }
                Op::f32x4_mul => { self.simd_f32x4_binop(|a,b| a*b)?; }
                Op::f32x4_div => { self.simd_f32x4_binop(|a,b| a/b)?; }
                Op::f32x4_extract_lane => { let lane = self.read_byte() as usize & 3; if let Value::V128(a) = self.pop() { let v = f32::from_le_bytes(a[lane*4..lane*4+4].try_into().unwrap()); self.push(Value::F64(v as f64))?; } else { self.push(Value::F64(0.0))?; } }
                Op::f32x4_replace_lane => { let lane = self.read_byte() as usize & 3; let val = self.pop().as_f64() as f32; if let Value::V128(mut a) = self.pop() { a[lane*4..lane*4+4].copy_from_slice(&val.to_le_bytes()); self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::i8x16_splat => { let v = self.pop().as_i32() as u8; self.push(Value::V128([v;16]))?; }
                Op::i8x16_add => { self.simd_i8x16_binop(|a,b| a.wrapping_add(b))?; }
                Op::i8x16_sub => { self.simd_i8x16_binop(|a,b| a.wrapping_sub(b))?; }
                Op::i8x16_eq =>  { self.simd_i8x16_binop(|a,b| if a==b {0xFF} else {0})?; }
                Op::i8x16_extract_lane_s => { let lane = self.read_byte() as usize & 15; if let Value::V128(a) = self.pop() { self.push(Value::I32(a[lane] as i8 as i32))?; } else { self.push(Value::I32(0))?; } }
                Op::i8x16_extract_lane_u => { let lane = self.read_byte() as usize & 15; if let Value::V128(a) = self.pop() { self.push(Value::I32(a[lane] as i32))?; } else { self.push(Value::I32(0))?; } }
                Op::i8x16_replace_lane => { let lane = self.read_byte() as usize & 15; let val = self.pop().as_i32() as u8; if let Value::V128(mut a) = self.pop() { a[lane] = val; self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::i8x16_shuffle => { let mut indices = [0u8;16]; for i in 0..16 { indices[i] = self.read_byte(); } let b = self.pop(); let a = self.pop(); if let (Value::V128(va), Value::V128(vb)) = (a,b) { let combined: Vec<u8> = va.iter().chain(vb.iter()).copied().collect(); let mut out = [0u8;16]; for i in 0..16 { out[i] = combined.get(indices[i] as usize).copied().unwrap_or(0); } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::i8x16_swizzle => { let b = self.pop(); let a = self.pop(); if let (Value::V128(va), Value::V128(vb)) = (a,b) { let mut out = [0u8;16]; for i in 0..16 { let idx = vb[i] as usize; out[i] = if idx < 16 { va[idx] } else { 0 }; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::i16x8_splat => { let v = self.pop().as_i32() as i16; let b = v.to_le_bytes(); let mut out = [0u8;16]; for i in 0..8 { out[i*2..i*2+2].copy_from_slice(&b); } self.push(Value::V128(out))?; }
                Op::i16x8_add => { self.simd_i16x8_binop(|a,b| a.wrapping_add(b))?; }
                Op::i16x8_sub => { self.simd_i16x8_binop(|a,b| a.wrapping_sub(b))?; }
                Op::i16x8_mul => { self.simd_i16x8_binop(|a,b| a.wrapping_mul(b))?; }
                Op::i16x8_extract_lane_s => { let lane = self.read_byte() as usize & 7; if let Value::V128(a) = self.pop() { let v = i16::from_le_bytes([a[lane*2], a[lane*2+1]]); self.push(Value::I32(v as i32))?; } else { self.push(Value::I32(0))?; } }
                Op::i16x8_extract_lane_u => { let lane = self.read_byte() as usize & 7; if let Value::V128(a) = self.pop() { let v = u16::from_le_bytes([a[lane*2], a[lane*2+1]]); self.push(Value::I32(v as i32))?; } else { self.push(Value::I32(0))?; } }
                Op::i16x8_replace_lane => { let lane = self.read_byte() as usize & 7; let val = self.pop().as_i32() as i16; if let Value::V128(mut a) = self.pop() { a[lane*2..lane*2+2].copy_from_slice(&val.to_le_bytes()); self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }

                // v128 bitwise
                Op::v128_and => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] & b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::v128_or  => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] | b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::v128_xor => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] ^ b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::v128_not => { if let Value::V128(a) = self.pop() { let mut out = [0u8;16]; for i in 0..16 { out[i] = !a[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::v128_andnot => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] & !b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                Op::v128_any_true => { if let Value::V128(a) = self.pop() { self.push(Value::I32(if a.iter().any(|&b| b != 0) { 1 } else { 0 }))?; } else { self.push(Value::I32(0))?; } }
                Op::v128_bitselect => { let mask = self.pop(); let v2 = self.pop(); let v1 = self.pop(); if let (Value::V128(a), Value::V128(b), Value::V128(m)) = (v1, v2, mask) { let mut out = [0u8;16]; for i in 0..16 { out[i] = (a[i] & m[i]) | (b[i] & !m[i]); } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }

                // -- Atomics (single-threaded: same as non-atomic for now) --
                Op::atomic_fence => {} // no-op in single-threaded
                Op::i32_atomic_load => { let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_load_i32(addr)))?; }
                // Atomic store: stack is [addr, value] — pop value first (top), then addr
                Op::i32_atomic_store => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.memory.atomic_store_i32(addr, v); }
                // Atomic RMW: stack is [addr, value] — pop value (top), addr (second), return old
                Op::i32_atomic_rmw_add => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_add_i32(addr, v)))?; }
                Op::i32_atomic_rmw_sub => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_sub_i32(addr, v)))?; }
                Op::i32_atomic_rmw_and => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_and_i32(addr, v)))?; }
                Op::i32_atomic_rmw_or  => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_or_i32(addr, v)))?; }
                Op::i32_atomic_rmw_xor => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_xor_i32(addr, v)))?; }
                Op::i32_atomic_rmw_xchg => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_xchg_i32(addr, v)))?; }
                // Atomic CmpXchg: stack is [addr, expected, replacement]
                Op::i32_atomic_rmw_cmpxchg => { let replacement = self.pop().as_i32(); let expected = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_cmpxchg_i32(addr, expected, replacement)))?; }
                Op::i64_atomic_load => { let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_load_i64(addr)))?; }
                Op::i64_atomic_store => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.memory.atomic_store_i64(addr, v); }
                Op::i64_atomic_rmw_add => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_rmw_add_i64(addr, v)))?; }
                Op::i64_atomic_rmw_sub => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_rmw_sub_i64(addr, v)))?; }
                Op::i64_atomic_rmw_cmpxchg => { let repl = self.pop().as_i64(); let exp = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_cmpxchg_i64(addr, exp, repl)))?; }
                Op::memory_atomic_wait32 => { let timeout = self.pop().as_i64(); let expected = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.wait32(addr, expected, timeout)))?; }
                Op::memory_atomic_notify => { let count = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.notify(addr, count)))?; }

                // -- Memory64 --
                Op::i64_memory_size => { self.push(Value::I64((self.memory.len() / 65536) as i64))?; }
                Op::i64_memory_grow => { let pages = self.pop().as_i64() as usize; let old = self.memory.grow(pages); self.push(Value::I64(old as i64))?; }
                Op::i32_load_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::I32(self.memory.load_i32(addr)?))?; }
                Op::i64_load_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::I64(self.memory.load_i64(addr)?))?; }
                Op::f64_load_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::F64(self.memory.load_f64(addr)?))?; }
                Op::i32_store_64 => { let v = self.pop().as_i32(); let addr = self.pop().as_i64() as usize; self.memory.store_i32(addr, v)?; }
                Op::i64_store_64 => { let v = self.pop().as_i64(); let addr = self.pop().as_i64() as usize; let _ = self.memory.store_i64(addr, v); }
                Op::f64_store_64 => { let v = self.pop().as_f64(); let addr = self.pop().as_i64() as usize; let _ = self.memory.store_f64(addr, v); }

                // -- Relaxed SIMD FMA --
                Op::f32x4_relaxed_madd => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i*4..i*4+4].try_into().unwrap());
                            out[i*4..i*4+4].copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f32x4_relaxed_nmadd => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i*4..i*4+4].try_into().unwrap());
                            out[i*4..i*4+4].copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f64x2_relaxed_madd => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                Op::f64x2_relaxed_nmadd => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // -- JS Promise Integration (JSPI) --
                Op::promise_suspend => {
                    // Explicit JSPI suspend point. Like await, but can be inserted
                    // by the compiler for synchronous-looking code that calls async APIs.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        let is_pending = o.properties.get("__type")
                            .map(|v| format!("{}", v) == "Promise")
                            .unwrap_or(false)
                            && o.properties.get("__state")
                                .map(|v| format!("{}", v) == "pending")
                                .unwrap_or(false);
                        if is_pending {
                            let promise_id = o.properties.get("__id")
                                .map(|v| v.as_f64() as u64)
                                .unwrap_or(0);
                            drop(o);
                            let fiber = self.save_fiber();
                            self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                            return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                        }
                        let resolved = o.properties.get("__value").cloned().unwrap_or(Value::Null);
                        drop(o);
                        self.push(resolved)?;
                    } else {
                        // Not a promise — return value as-is
                        self.push(val)?;
                    }
                }

                // -- WASM GC Type System --
                Op::set_type_id => {
                    // Stack: [obj, type_id_i32] → [obj]
                    let type_id = self.pop().as_i32() as usize;
                    let obj = self.peek(0);
                    if let Value::Object(o) = obj {
                        o.lock().unwrap().type_id = type_id;
                    }
                }

                // -- Iteration protocol --
                Op::iter_get => {
                    // Get an iterator from an iterable.
                    // Checks: __iter__ (Python), Symbol.iterator (JS), GetEnumerator (C#)
                    // Fallback: if it's an array, create an index-based iterator object.
                    let iterable = self.pop();
                    let mut iter_obj = Object::new();

                    match &iterable {
                        Value::Object(o) => {
                            let ob = o.lock().unwrap();
                            // Check for __iter__ method
                            if ob.properties.contains_key("__iter__") {
                                // Store the iterable and we'll call __iter__ via iter_next
                                drop(ob);
                                iter_obj.properties.insert("__iterable".into(), iterable.clone());
                                iter_obj.properties.insert("__protocol".into(), Value::String(Arc::from("dunder")));
                                iter_obj.properties.insert("__started".into(), Value::Bool(false));
                            } else if matches!(&ob.kind, ObjectKind::Array(_)) {
                                // Array: index-based iteration
                                drop(ob);
                                iter_obj.properties.insert("__iterable".into(), iterable.clone());
                                iter_obj.properties.insert("__protocol".into(), Value::String(Arc::from("array")));
                                iter_obj.properties.insert("__index".into(), Value::I32(0));
                            } else {
                                // Dict/Object: iterate keys
                                let keys: Vec<Value> = ob.properties.keys()
                                    .filter(|k| !k.starts_with("__"))
                                    .map(|k| Value::String(Arc::from(k.as_str())))
                                    .collect();
                                drop(ob);
                                let keys_arr = Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                                iter_obj.properties.insert("__iterable".into(), keys_arr);
                                iter_obj.properties.insert("__protocol".into(), Value::String(Arc::from("array")));
                                iter_obj.properties.insert("__index".into(), Value::I32(0));
                            }
                        }
                        Value::String(s) => {
                            // String: iterate characters
                            let chars: Vec<Value> = s.chars()
                                .map(|c| Value::String(Arc::from(c.to_string().as_str())))
                                .collect();
                            let chars_arr = Value::Object(Arc::new(Mutex::new(Object::new_array(chars))));
                            iter_obj.properties.insert("__iterable".into(), chars_arr);
                            iter_obj.properties.insert("__protocol".into(), Value::String(Arc::from("array")));
                            iter_obj.properties.insert("__index".into(), Value::I32(0));
                        }
                        _ => {
                            // Not iterable
                            iter_obj.properties.insert("__protocol".into(), Value::String(Arc::from("empty")));
                        }
                    }
                    iter_obj.properties.insert("__type".into(), Value::String(Arc::from("iterator")));
                    iter_obj.properties.insert("__done".into(), Value::Bool(false));
                    self.push(Value::Object(Arc::new(Mutex::new(iter_obj))))?;
                }
                Op::iter_next => {
                    // Advance iterator. Returns {value, done}.
                    // Stack: [iterator] → [value, bool_done]
                    let iter_val = self.pop();
                    if let Value::Object(ref iter_obj) = iter_val {
                        let protocol = {
                            let ob = iter_obj.lock().unwrap();
                            ob.properties.get("__protocol").map(|v| v.to_string()).unwrap_or_default()
                        };

                        match protocol.as_str() {
                            "array" => {
                                let (value, done) = {
                                    let mut ob = iter_obj.lock().unwrap();
                                    let idx = ob.properties.get("__index")
                                        .map(|v| v.as_i32() as usize).unwrap_or(0);
                                    let iterable = ob.properties.get("__iterable").cloned().unwrap_or(Value::Null);
                                    let (val, is_done) = if let Value::Object(arr) = &iterable {
                                        let arr_ob = arr.lock().unwrap();
                                        if let ObjectKind::Array(elems) = &arr_ob.kind {
                                            if idx < elems.len() {
                                                (elems[idx].clone(), false)
                                            } else {
                                                (Value::Null, true)
                                            }
                                        } else {
                                            (Value::Null, true)
                                        }
                                    } else {
                                        (Value::Null, true)
                                    };
                                    ob.properties.insert("__index".into(), Value::I32((idx + 1) as i32));
                                    ob.properties.insert("__done".into(), Value::Bool(is_done));
                                    (val, is_done)
                                };
                                self.push(value)?;
                                self.push(Value::Bool(done))?;
                            }
                            "dunder" => {
                                // Python __iter__/__next__ protocol
                                // First call: call __iter__ to get the iterator object
                                // Subsequent: call __next__ on the iterator
                                let started = {
                                    let ob = iter_obj.lock().unwrap();
                                    ob.properties.get("__started").map(|v| dyn_truthy(v)).unwrap_or(false)
                                };
                                if !started {
                                    // Call __iter__ on the iterable
                                    let iterable = iter_obj.lock().unwrap().properties.get("__iterable").cloned().unwrap_or(Value::Null);
                                    if let Value::Object(ref it_obj) = iterable {
                                        let iter_fn = it_obj.lock().unwrap().properties.get("__iter__").cloned();
                                        if let Some(func) = iter_fn {
                                            self.push(func)?;
                                            self.push(iterable.clone())?;
                                            self.call_value(1)?;
                                            // __iter__ returns the iterator — store it
                                            // For simplicity, assume __iter__ returns self or an array
                                        }
                                    }
                                    iter_obj.lock().unwrap().properties.insert("__started".into(), Value::Bool(true));
                                }
                                // For now, treat dunder iterators as done (full protocol needs coroutines)
                                self.push(Value::Null)?;
                                self.push(Value::Bool(true))?;
                            }
                            _ => {
                                self.push(Value::Null)?;
                                self.push(Value::Bool(true))?;
                            }
                        }
                    } else {
                        self.push(Value::Null)?;
                        self.push(Value::Bool(true))?;
                    }
                }
                Op::spread => {
                    // Spread array onto stack: [array] → [elem0, elem1, ...]
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(elems) = &o.kind {
                            for elem in elems {
                                self.push(elem.clone())?;
                            }
                        }
                    }
                }
                Op::class_new => { let _ = self.read_u16(); self.push(Value::Null)?; }
                Op::method_def => { let _ = self.read_u16(); }
                Op::inherit => { self.pop(); }
            }
        }
    }

    /// Try to call a dunder method (__add__, __lt__, etc.) on an object.
    /// Returns Some(result) if the method exists, None otherwise.
    fn try_dunder_binary(&mut self, obj: &Arc<Mutex<crate::value::Object>>, arg: &Value, dunder: &str) -> Option<Value> {
        let method = {
            let o = obj.lock().unwrap();
            o.properties.get(dunder).cloned()
        };
        if let Some(func_val) = method {
            // Call dunder(self, arg) — push func, self, arg, call(2)
            let self_val = Value::Object(obj.clone());
            self.push(func_val).ok()?;
            self.push(self_val).ok()?;
            self.push(arg.clone()).ok()?;
            self.call_value(2).ok()?;
            // Execute until the function returns
            self.execute_until(self.frames.len()).ok()?;
            Some(self.pop())
        } else {
            None
        }
    }

    fn call_value(&mut self, argc: usize) -> Result<(), VMError> {
        let callee_idx = self.stack.len() - 1 - argc;
        let callee = self.stack[callee_idx].clone();

        match &callee {
            Value::Object(obj) => {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Function(func) => {
                        self.call_function(func, argc)?;
                    }
                    ObjectKind::HostFunction(idx) => {
                        let idx = *idx;
                        drop(o);
                        let args: Vec<Value> = self.stack[self.stack.len() - argc..].to_vec();
                        for _ in 0..argc { self.stack.pop(); }
                        self.stack.pop();
                        let placeholder: HostFn = Arc::new(|_, _| Value::Null);
                        let host_fn = std::mem::replace(&mut self.host_fns[idx], placeholder);
                        let result = {
                            let mut ctx = self.make_host_context();
                            host_fn(&mut ctx, &args)
                        };
                        self.host_fns[idx] = host_fn;
                        self.push(result)?;
                    }
                    _other => {
                        // Check for __call__ dunder (Python callable objects)
                        let call_fn = o.properties.get("__call__").cloned();
                        let kind_name = format!("{:?}", std::mem::discriminant(&o.kind));
                        drop(o);
                        if let Some(func) = call_fn {
                            self.stack[callee_idx] = func;
                            return self.call_value(argc);
                        }
                        let chunk_name = if !self.frames.is_empty() {
                            self.chunks[self.frame().chunk_index].name.clone()
                        } else { "?".into() };
                        return Err(VMError::new(format!("Not a function in chunk '{}' (kind: {})",
                            chunk_name, kind_name)));
                    }
                }
            }
            _ => {
                let stack = self.capture_call_stack();
                return Err(VMError::new(format!("{} is not callable (type: {})", callee.type_tag(), callee)).with_stack(stack));
            }
        }
        Ok(())
    }

    fn call_function(&mut self, func: &Function, argc: usize) -> Result<(), VMError> {
        if self.frames.len() >= MAX_FRAMES {
            return Err(VMError::new("Stack overflow"));
        }

        let chunk_index = func.chunk_index;
        let arity = func.arity as usize;
        let base = self.stack.len() - argc - 1;


        // Arity validation: warn on mismatch but don't trap.
        // Dynamic languages (JS, Python, Ruby) rely on flexible arity.
        // WASM strict mode would trap here; we pad missing args with Null
        // and ignore excess args (they stay on caller's stack frame).
        if argc > arity && arity > 0 {
            // Excess arguments — truncate to expected arity.
            // Pop the extras so they don't corrupt the callee's locals.
            for _ in 0..(argc - arity) {
                self.pop();
            }
        }
        for _ in argc..arity {
            self.push(Value::Null)?;
        }

        let local_count = self.chunks[chunk_index].local_count as usize;
        let total = 1 + local_count.max(arity);
        let have = self.stack.len() - base;
        for _ in have..total {
            self.push(Value::Null)?;
        }

        let upvalues = func.upvalues.clone();
        self.frames.push(CallFrame { chunk_index, ip: 0, base, upvalues });
        Ok(())
    }

    fn capture_upvalue(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        for uv in &self.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                if idx == stack_idx { return uv.clone(); }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue { location: UpvalueLocation::Open(stack_idx) }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    fn close_upvalues(&mut self, from: usize) {
        let mut i = 0;
        while i < self.open_upvalues.len() {
            let should_close = matches!(
                self.open_upvalues[i].lock().unwrap().location,
                UpvalueLocation::Open(idx) if idx >= from
            );
            if should_close {
                let uv = self.open_upvalues.remove(i);
                let mut u = uv.lock().unwrap();
                if let UpvalueLocation::Open(idx) = u.location {
                    u.location = UpvalueLocation::Closed(self.stack[idx].clone());
                }
            } else {
                i += 1;
            }
        }
    }
}

impl VM {
    /// WASM GC-style property/method resolution.
    ///
    /// Resolution order:
    /// 1. Property getter (__get_{name})
    /// 2. Instance property (on the object itself)
    /// 3. TypeRegistry vtable (type_id → method table → parent chain)
    /// 4. Legacy type_methods table (fallback for old code)
    /// 5. Universal Object methods (type 0)
    pub fn resolve_property(&self, obj: &Value, name: &str) -> Result<Value, VMError> {
        match obj {
            Value::Object(o) => {
                let ob = o.lock().unwrap();

                // 1. Instance property (getters handled in struct_get opcode directly)
                let val = ob.get(name);
                if !matches!(val, Value::Null) {
                    return Ok(val);
                }

                // 3. TypeRegistry vtable
                let type_id = ob.type_id;
                drop(ob); // release borrow before accessing self

                if type_id > 0 {
                    if let Some(method) = self.type_registry.resolve_method(type_id, name) {
                        return Ok(self.method_to_value(method));
                    }
                }

                // Also try inferring type from ObjectKind or __type property
                let ob = o.lock().unwrap();
                let inferred_type = ob.properties.get("__type")
                    .map(|v| format!("{}", v).to_lowercase())
                    .unwrap_or_else(|| match &ob.kind {
                        ObjectKind::Array(_) => "list".into(),
                        _ => String::new(),
                    });
                drop(ob);

                if !inferred_type.is_empty() {
                    if let Some(tid) = self.type_registry.get_id(&inferred_type) {
                        if let Some(method) = self.type_registry.resolve_method(tid, name) {
                            return Ok(self.method_to_value(method));
                        }
                    }
                }

                // 3. Universal Object methods (type 0)
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }

                Ok(Value::Null)
            }
            Value::String(s) => {
                if name == "length" {
                    return Ok(Value::F64(s.len() as f64));
                }
                if let Some(tid) = self.type_registry.get_id("string") {
                    if let Some(method) = self.type_registry.resolve_method(tid, name) {
                        return Ok(self.method_to_value(method));
                    }
                }
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Null)
            }
            _ => {
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Null)
            }
        }
    }

    /// Convert a Method (from TypeRegistry) to a callable Value.
    /// Uses the function table for zero-allocation dispatch.
    fn method_to_value(&self, method: &crate::typedef::Method) -> Value {
        match method {
            crate::typedef::Method::HostFn(idx) => {
                // Return existing entry from function table — no allocation
                if *idx < self.func_table.len() {
                    self.func_table[*idx].clone()
                } else {
                    // Fallback: create new (shouldn't happen if registered properly)
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::HostFunction(*idx);
                    Value::Object(Arc::new(Mutex::new(obj)))
                }
            }
            crate::typedef::Method::ChunkFn(idx) => {
                let chunk = &self.chunks[*idx];
                let func = Function {
                    name: Some(chunk.name.clone()),
                    arity: chunk.arity,
                    chunk_index: *idx,
                    upvalues: Vec::new(),
                };
                let obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0, fields: Vec::new() };
                Value::Object(Arc::new(Mutex::new(obj)))
            }
        }
    }
}

fn dyn_truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        Value::Object(o) => {
            let ob = o.lock().unwrap();
            // Check __bool__ property (set by Python classes as a bool value).
            // If __bool__ is a Bool value, use it. If it's a function, we can't
            // call it from here (no VM access) — the compiler should call it
            // and set the result as a bool property.
            if let Some(Value::Bool(b)) = ob.properties.get("__bool__") {
                return *b;
            }
            // JS semantics: all objects are truthy (including empty arrays/objects).
            // Python's bool() builtin handles empty-container checks at the compiler level.
            true
        }
        Value::WeakRef(w) => w.upgrade().is_some(),
        Value::V128(b) => b.iter().any(|&x| x != 0),
    }
}
