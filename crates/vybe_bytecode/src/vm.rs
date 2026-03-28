use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::{EventLoop, Task};
use crate::fiber::{Fiber, SavedFrame};
use crate::opcode::Op;
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

/// Host function signature. Receives args, returns a value.
pub type HostFn = Box<dyn Fn(&[Value]) -> Value>;

#[derive(Debug, Clone)]
struct CallFrame {
    chunk_index: usize,
    ip: usize,
    base: usize,
    upvalues: Vec<Rc<RefCell<Upvalue>>>,
}

/// Exception handler entry — pushed by try_start, popped by try_end or catch.
#[derive(Debug, Clone)]
struct ExceptionHandler {
    /// Instruction pointer to jump to on catch.
    catch_ip: usize,
    /// Chunk index the handler was registered in.
    chunk_index: usize,
    /// Stack depth when try_start was executed (for unwinding).
    stack_depth: usize,
    /// Call frame depth when try_start was executed.
    frame_depth: usize,
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
    open_upvalues: Vec<Rc<RefCell<Upvalue>>>,
    host_fns: Vec<HostFn>,
    /// Registry: (module, name) → index into host_fns.
    pub host_registry: HashMap<(String, String), usize>,
    /// Import resolution table: import_index → host_fn_index.
    import_table: Vec<usize>,
    /// Exception handler stack (WASM exception proposal).
    exception_handlers: Vec<ExceptionHandler>,
    /// Event loop for async operations (shared with host functions).
    pub event_loop: Rc<RefCell<EventLoop>>,
    /// WASM GC-style type definitions with vtable method dispatch.
    pub type_registry: crate::typedef::TypeRegistry,
    /// Linear memory (WASM MVP) — byte buffer for binary data.
    pub memory: Vec<u8>,
    /// Function table (WASM MVP) — for call_indirect.
    pub func_table: Vec<Value>,
    /// Block label stack for structured control flow.
    label_stack: Vec<LabelEntry>,
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
            import_table: Vec::new(),
            exception_handlers: Vec::new(),
            event_loop: Rc::new(RefCell::new(EventLoop::new())),
            type_registry: crate::typedef::TypeRegistry::new(),
            memory: Vec::new(),
            func_table: Vec::new(),
            label_stack: Vec::new(),
        }
    }

    /// Register a host function with a (module, name) pair.
    /// Also adds it to the function table for call_indirect dispatch.
    pub fn register_host_fn(&mut self, module: &str, name: &str, f: HostFn) {
        let idx = self.host_fns.len();
        self.host_fns.push(f);
        self.host_registry.insert((module.to_string(), name.to_string()), idx);
        // Add to function table — func_table index == host_fns index for host functions
        while self.func_table.len() <= idx {
            self.func_table.push(Value::Null);
        }
        // Store as a lightweight marker — call_indirect will recognize host fn indices
        let mut obj = Object::new();
        obj.kind = ObjectKind::HostFunction(idx);
        self.func_table[idx] = Value::Object(Rc::new(RefCell::new(obj)));
    }

    /// Get a type_id by name from the TypeRegistry.
    pub fn get_type_id(&self, name: &str) -> usize {
        self.type_registry.get_id(name).unwrap_or(0)
    }


    /// Load chunks and execute the script chunk (first in the new set).
    /// Appends to existing chunks so cross-language calls work (functions reference chunk indices).
    /// Resolves the import table against registered host functions.
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

        // Resolve imports from the new script chunk's import table
        self.import_table.clear();
        for import in &self.chunks[script_idx].imports {
            let key = (import.module.clone(), import.name.clone());
            match self.host_registry.get(&key) {
                Some(&idx) => self.import_table.push(idx),
                None => return Err(VMError::new(format!(
                    "Unresolved import: \"{}\" \"{}\"", import.module, import.name
                ))),
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
    fn test_type(&self, val: &Value, target_name: &str) -> bool {
        match val {
            Value::Object(o) => {
                let ob = o.borrow();
                if ob.type_id > 0 {
                    if let Some(target_id) = self.type_registry.get_id(target_name) {
                        self.type_registry.is_subtype(ob.type_id, target_id)
                    } else { false }
                } else {
                    let obj_type = ob.properties.get("__type")
                        .map(|v| format!("{}", v).to_lowercase())
                        .or_else(|| ob.properties.get("__control_type")
                            .map(|v| format!("{}", v).to_lowercase()))
                        .unwrap_or_default();
                    if obj_type == target_name { true }
                    else if let Some(tid) = self.type_registry.get_id(&obj_type) {
                        if let Some(target_id) = self.type_registry.get_id(target_name) {
                            self.type_registry.is_subtype(tid, target_id)
                        } else { false }
                    } else {
                        target_name == "object"
                    }
                }
            }
            Value::String(_) => target_name == "string" || target_name == "object",
            Value::F64(_) | Value::I32(_) | Value::I64(_) => {
                target_name == "integer" || target_name == "double" || target_name == "number" || target_name == "object"
            }
            Value::Bool(_) => target_name == "boolean" || target_name == "object",
            Value::V128(_) => target_name == "v128",
            Value::Null | Value::Undefined => false,
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
        loop {
            let f = self.frame();
            let chunk = &self.chunks[f.chunk_index];

            if f.ip >= chunk.code.len() {
                if self.frames.len() <= 1 {
                    return Ok(Value::Null);
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

            match op {
                Op::halt => {
                    // Close all open upvalues so closures retain captured values
                    self.close_upvalues(0);
                    return Ok(if self.stack.is_empty() { Value::Null } else { self.pop() });
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
                    let val = self.globals.get(&name).cloned().unwrap_or(Value::Undefined);
                    self.push(val)?;
                }
                Op::global_set => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.peek(0).clone();
                    self.globals.insert(name, val);
                }
                Op::upvalue_get => {
                    let idx = self.read_byte() as usize;
                    let uv = self.frame().upvalues[idx].clone();
                    let val = match &uv.borrow().location {
                        UpvalueLocation::Open(si) => self.stack[*si].clone(),
                        UpvalueLocation::Closed(v) => v.clone(),
                    };
                    self.push(val)?;
                }
                Op::upvalue_set => {
                    let idx = self.read_byte() as usize;
                    let val = self.peek(0).clone();
                    let uv = self.frame().upvalues[idx].clone();
                    let mut u = uv.borrow_mut();
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
                    // Check for getter first — needs auto-invoke (can't do in resolve_property
                    // because it needs mutable self for call_value)
                    if let Value::Object(ref o) = obj {
                        let getter_key = format!("__get_{}", name);
                        let getter = o.borrow().properties.get(&getter_key).cloned();
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
                        let setter = o.borrow().properties.get(&setter_key).cloned();
                        if let Some(setter_fn) = setter {
                            // Call the setter with this = obj, value = val
                            self.push(setter_fn)?;
                            self.push(obj.clone())?;
                            self.push(val.clone())?;
                            self.call_value(2)?;
                            // Setter returns are discarded, push the assigned value
                            self.pop(); // discard setter return
                            self.push(val)?;
                        } else {
                            o.borrow_mut().set(name, val.clone());
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
                            let k = format!("{}", key);
                            let val = o.borrow().get(&k);
                            self.push(val)?;
                        }
                        Value::String(s) => {
                            let i = key.as_f64() as usize;
                            if let Some(ch) = s.chars().nth(i) {
                                self.push(Value::String(Rc::from(ch.to_string().as_str())))?;
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
                        let k = format!("{}", key);
                        o.borrow_mut().set(k, val.clone());
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
                    self.push(Value::I32(if b == 0 { 0 } else { a.wrapping_div(b) }))?;
                }
                Op::i32_div_u => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(if b == 0 { 0 } else { (a / b) as i32 }))?;
                }
                Op::i32_rem_s => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(if b == 0 { 0 } else { a.wrapping_rem(b) }))?;
                }
                Op::i32_rem_u => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(if b == 0 { 0 } else { (a % b) as i32 }))?;
                }

                // -- i64 arithmetic --
                Op::i64_add => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_add(b)))?; }
                Op::i64_sub => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_sub(b)))?; }
                Op::i64_mul => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_mul(b)))?; }
                Op::i64_div_s => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(if b == 0 { 0 } else { a.wrapping_div(b) }))?; }
                Op::i64_div_u => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(if b == 0 { 0 } else { (a / b) as i64 }))?; }
                Op::i64_rem_s => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(if b == 0 { 0 } else { a.wrapping_rem(b) }))?; }
                Op::i64_rem_u => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(if b == 0 { 0 } else { (a % b) as i64 }))?; }
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
                    self.push(Value::String(Rc::from(s.as_str())))?;
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
                    self.push(Value::String(Rc::from(result.as_str())))?;
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
                    if self.frames.is_empty() {
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
                    let mut upvalues: Vec<Rc<RefCell<Upvalue>>> = Vec::with_capacity(uv_count);
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
                    let obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0 };
                    self.push(Value::Object(Rc::new(RefCell::new(obj))))?;
                }

                // -- Host functions --
                Op::call_import => {
                    let import_idx = self.read_u16() as usize;
                    let argc = self.read_byte() as usize;
                    let base = self.stack.len() - argc;
                    let args: Vec<Value> = self.stack[base..].to_vec();
                    self.stack.truncate(base);

                    if import_idx < self.import_table.len() {
                        let host_idx = self.import_table[import_idx];
                        let result = (self.host_fns[host_idx])(&args);

                        // JSPI: if host function returned a pending Promise,
                        // transparently suspend and resume when resolved.
                        // The calling code doesn't need `await` — it looks synchronous.
                        if let Value::Object(ref obj) = result {
                            let o = obj.borrow();
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
                                // JSPI suspend: save entire VM state as a fiber
                                let fiber = self.save_fiber();
                                self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                                return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                            }
                        }

                        self.push(result)?;
                    } else {
                        return Err(VMError::new(format!("Unresolved import index: {}", import_idx)));
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
                    self.push(Value::Object(Rc::new(RefCell::new(obj))))?;
                }
                Op::array_new => {
                    let count = self.read_u16() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(elems)))))?;
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
                // ref_test: TypeOf...Is using TypeRegistry hierarchy
                Op::ref_test => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.pop();
                    let result = match &val {
                        Value::Object(o) => {
                            let ob = o.borrow();
                            // Check type_id via TypeRegistry
                            if ob.type_id > 0 {
                                if let Some(target_id) = self.type_registry.get_id(&target_name) {
                                    self.type_registry.is_subtype(ob.type_id, target_id)
                                } else {
                                    false
                                }
                            } else {
                                // Fallback: check __type property name match
                                let obj_type = ob.properties.get("__type")
                                    .map(|v| format!("{}", v).to_lowercase())
                                    .or_else(|| ob.properties.get("__control_type")
                                        .map(|v| format!("{}", v).to_lowercase()))
                                    .unwrap_or_default();
                                if obj_type == target_name { true }
                                else if let Some(tid) = self.type_registry.get_id(&obj_type) {
                                    if let Some(target_id) = self.type_registry.get_id(&target_name) {
                                        self.type_registry.is_subtype(tid, target_id)
                                    } else { false }
                                } else {
                                    target_name == "object" // everything is Object
                                }
                            }
                        }
                        Value::String(_) => target_name == "string" || target_name == "object",
                        Value::F64(_) | Value::I32(_) | Value::I64(_) => {
                            target_name == "integer" || target_name == "double" || target_name == "object"
                        }
                        Value::Bool(_) => target_name == "boolean" || target_name == "object",
                        Value::V128(_) => target_name == "v128",
                        Value::Null | Value::Undefined => false,
                    };
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
                    let is_fn = matches!(&v, Value::Object(o) if matches!(o.borrow().kind, ObjectKind::Function(_)));
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
                            Value::String(Rc::from(format!("{}{}", a, b).as_str()))
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
                        (Value::Object(x), Value::Object(y)) => Rc::ptr_eq(x, y),
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
                        (Value::Object(x), Value::Object(y)) => !Rc::ptr_eq(x, y),
                        _ => true,
                    };
                    self.push(Value::Bool(result))?;
                }
                Op::dyn_lt => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x < *y,
                        _ => a.as_f64() < b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_gt => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x > *y,
                        _ => a.as_f64() > b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_le => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x <= *y,
                        _ => a.as_f64() <= b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                Op::dyn_ge => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x >= *y,
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
                        let o = obj.borrow();
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
                        chunk_index: f.chunk_index,
                        stack_depth: self.stack.len(),
                        frame_depth: self.frames.len(),
                    });
                }
                Op::try_end => {
                    // Normal exit from try block — pop the handler
                    self.exception_handlers.pop();
                }
                Op::throw => {
                    let val = self.pop();
                    if let Some(handler) = self.exception_handlers.pop() {
                        // Unwind: restore stack and frames to the state at try_start
                        while self.frames.len() > handler.frame_depth {
                            let base = self.frames.last().unwrap().base;
                            self.close_upvalues(base);
                            self.frames.pop();
                        }
                        self.stack.truncate(handler.stack_depth);
                        // Push the exception value (for catch binding)
                        self.push(val)?;
                        // Jump to catch block
                        let f = self.frame_mut();
                        f.ip = handler.catch_ip;
                    } else {
                        // No handler — propagate as VM error
                        return Err(VMError::new(format!("{}", val)));
                    }
                }

                Op::throw_ref => {
                    // Same as throw — value is already a reference
                    let val = self.pop();
                    if let Some(handler) = self.exception_handlers.pop() {
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
                        return Err(VMError::new(format!("{}", val)));
                    }
                }
                Op::try_table => {
                    // Modern block-based try: [try_table, u8 handler_count, ...]
                    let handler_count = self.read_byte() as usize;
                    for _ in 0..handler_count {
                        let _tag = self.read_byte();
                        let offset = self.read_u16();
                        // Register handler at offset
                        let ip = self.frame().ip + offset as usize;
                        self.exception_handlers.push(ExceptionHandler {
                            catch_ip: ip,
                            stack_depth: self.stack.len(),
                            frame_depth: self.frames.len(),
                            chunk_index: self.frame().chunk_index,
                        });
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
                    let pages = (self.memory.len() / 65536) as i32;
                    self.push(Value::I32(pages))?;
                }
                Op::memory_grow => {
                    let pages = self.pop().as_f64() as usize;
                    let old_pages = self.memory.len() / 65536;
                    self.memory.resize(self.memory.len() + pages * 65536, 0);
                    self.push(Value::I32(old_pages as i32))?;
                }
                Op::i32_load => {
                    let addr = self.pop().as_f64() as usize;
                    if addr + 4 <= self.memory.len() {
                        let val = i32::from_le_bytes([self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3]]);
                        self.push(Value::I32(val))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::i32_store => {
                    let val = self.pop().as_f64() as i32;
                    let addr = self.pop().as_f64() as usize;
                    if addr + 4 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+4].copy_from_slice(&bytes);
                    }
                }
                Op::i64_load => {
                    let addr = self.pop().as_f64() as usize;
                    if addr + 8 <= self.memory.len() {
                        let val = i64::from_le_bytes([
                            self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3],
                            self.memory[addr+4], self.memory[addr+5], self.memory[addr+6], self.memory[addr+7],
                        ]);
                        self.push(Value::I64(val))?;
                    } else {
                        self.push(Value::I64(0))?;
                    }
                }
                Op::i64_store => {
                    let val = self.pop().as_f64() as i64;
                    let addr = self.pop().as_f64() as usize;
                    if addr + 8 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+8].copy_from_slice(&bytes);
                    }
                }
                Op::f64_load => {
                    let addr = self.pop().as_f64() as usize;
                    if addr + 8 <= self.memory.len() {
                        let val = f64::from_le_bytes([
                            self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3],
                            self.memory[addr+4], self.memory[addr+5], self.memory[addr+6], self.memory[addr+7],
                        ]);
                        self.push(Value::F64(val))?;
                    } else {
                        self.push(Value::F64(0.0))?;
                    }
                }
                Op::f64_store => {
                    let val = self.pop().as_f64();
                    let addr = self.pop().as_f64() as usize;
                    if addr + 8 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+8].copy_from_slice(&bytes);
                    }
                }
                Op::i32_load8_u => {
                    let addr = self.pop().as_f64() as usize;
                    if addr < self.memory.len() {
                        self.push(Value::I32(self.memory[addr] as i32))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::i32_store8 => {
                    let val = self.pop().as_f64() as u8;
                    let addr = self.pop().as_f64() as usize;
                    if addr < self.memory.len() {
                        self.memory[addr] = val;
                    }
                }
                Op::f32_load => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 4 <= self.memory.len() {
                        let val = f32::from_le_bytes([self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3]]);
                        self.push(Value::F64(val as f64))?;
                    } else { self.push(Value::F64(0.0))?; }
                }
                Op::f32_store => {
                    let val = self.pop().as_f64() as f32;
                    let addr = self.pop().as_i32() as usize;
                    if addr + 4 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+4].copy_from_slice(&bytes);
                    }
                }
                Op::i32_load8_s => {
                    let addr = self.pop().as_i32() as usize;
                    if addr < self.memory.len() {
                        self.push(Value::I32(self.memory[addr] as i8 as i32))?;
                    } else { self.push(Value::I32(0))?; }
                }
                Op::i32_load16_s => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let val = i16::from_le_bytes([self.memory[addr], self.memory[addr+1]]);
                        self.push(Value::I32(val as i32))?;
                    } else { self.push(Value::I32(0))?; }
                }
                Op::i32_load16_u => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let val = u16::from_le_bytes([self.memory[addr], self.memory[addr+1]]);
                        self.push(Value::I32(val as i32))?;
                    } else { self.push(Value::I32(0))?; }
                }
                Op::i32_store16 => {
                    let val = self.pop().as_i32() as i16;
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+2].copy_from_slice(&bytes);
                    }
                }
                Op::i64_load8_s => {
                    let addr = self.pop().as_i32() as usize;
                    if addr < self.memory.len() { self.push(Value::I64(self.memory[addr] as i8 as i64))?; }
                    else { self.push(Value::I64(0))?; }
                }
                Op::i64_load8_u => {
                    let addr = self.pop().as_i32() as usize;
                    if addr < self.memory.len() { self.push(Value::I64(self.memory[addr] as i64))?; }
                    else { self.push(Value::I64(0))?; }
                }
                Op::i64_load16_s => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let val = i16::from_le_bytes([self.memory[addr], self.memory[addr+1]]);
                        self.push(Value::I64(val as i64))?;
                    } else { self.push(Value::I64(0))?; }
                }
                Op::i64_load16_u => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let val = u16::from_le_bytes([self.memory[addr], self.memory[addr+1]]);
                        self.push(Value::I64(val as i64))?;
                    } else { self.push(Value::I64(0))?; }
                }
                Op::i64_load32_s => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 4 <= self.memory.len() {
                        let val = i32::from_le_bytes([self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3]]);
                        self.push(Value::I64(val as i64))?;
                    } else { self.push(Value::I64(0))?; }
                }
                Op::i64_load32_u => {
                    let addr = self.pop().as_i32() as usize;
                    if addr + 4 <= self.memory.len() {
                        let val = u32::from_le_bytes([self.memory[addr], self.memory[addr+1], self.memory[addr+2], self.memory[addr+3]]);
                        self.push(Value::I64(val as i64))?;
                    } else { self.push(Value::I64(0))?; }
                }
                Op::i64_store8 => {
                    let val = self.pop().as_i64() as u8;
                    let addr = self.pop().as_i32() as usize;
                    if addr < self.memory.len() { self.memory[addr] = val; }
                }
                Op::i64_store16 => {
                    let val = self.pop().as_i64() as i16;
                    let addr = self.pop().as_i32() as usize;
                    if addr + 2 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+2].copy_from_slice(&bytes);
                    }
                }
                Op::i64_store32 => {
                    let val = self.pop().as_i64() as i32;
                    let addr = self.pop().as_i32() as usize;
                    if addr + 4 <= self.memory.len() {
                        let bytes = val.to_le_bytes();
                        self.memory[addr..addr+4].copy_from_slice(&bytes);
                    }
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
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(values)))))?;
                }
                Op::unpack => {
                    let arr = self.pop();
                    if let Value::Object(obj) = arr {
                        let o = obj.borrow();
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
                        let ci = self.frame().chunk_index;
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
                    // Table index is on stack before args
                    let table_idx_pos = self.stack.len() - 1 - argc;
                    let table_idx = self.stack[table_idx_pos].as_f64() as usize;
                    if table_idx < self.func_table.len() {
                        self.stack[table_idx_pos] = self.func_table[table_idx].clone();
                        self.call_value(argc)?;
                    } else {
                        return Err(VMError::new(format!("call_indirect: table index {} out of bounds", table_idx)));
                    }
                }

                // -- Component Model (stubs for now) --
                Op::canon_lift => { let _ = self.read_u16(); }
                Op::canon_lower => { let _ = self.read_u16(); }

                // -- JS String Builtins (wasm:js-string proposal) --
                Op::str_length => {
                    let s = self.pop();
                    let len = match &s {
                        Value::String(s) => s.chars().count() as i32,
                        Value::Object(obj) => {
                            let o = obj.borrow();
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
                    self.push(Value::String(Rc::from(ch.to_string().as_str())))?;
                }
                Op::str_char_at => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let ch = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| Rc::from(c.to_string().as_str()))
                            .unwrap_or(Rc::from(""))
                    } else { Rc::from("") };
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
                        Rc::from(sub.as_str())
                    } else { Rc::from("") };
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
                        Rc::from(s.to_uppercase().as_str())
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_to_lower => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Rc::from(s.to_lowercase().as_str())
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Rc::from(s.trim()) } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim_start => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Rc::from(s.trim_start()) } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_trim_end => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Rc::from(s.trim_end()) } else { Rc::from("") };
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
                            Rc::from(s.replace(o.as_ref(), n.as_ref()).as_str())
                        }
                        _ => Rc::from(""),
                    };
                    self.push(Value::String(r))?;
                }
                Op::str_split => {
                    let delim = self.pop(); let s = self.pop();
                    let parts: Vec<Value> = match (&s, &delim) {
                        (Value::String(s), Value::String(d)) => {
                            s.split(d.as_ref()).map(|p| Value::String(Rc::from(p))).collect()
                        }
                        _ => vec![],
                    };
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(parts)))))?;
                }
                Op::str_repeat => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Rc::from(s.repeat(count).as_str())
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_pad_start => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Rc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Rc::from(format!("{}{}", fill_str, s).as_str())
                        }
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_pad_end => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Rc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Rc::from(format!("{}{}", s, fill_str).as_str())
                        }
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::str_reverse => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Rc::from(s.chars().rev().collect::<String>().as_str())
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                // Unicode code points (beyond BMP — emoji, CJK)
                Op::str_from_code_point => {
                    let cp = self.pop().as_i32() as u32;
                    let ch = char::from_u32(cp).unwrap_or('\u{FFFD}');
                    self.push(Value::String(Rc::from(ch.to_string().as_str())))?;
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
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(codes)))))?;
                }
                Op::str_from_char_codes => {
                    let arr = self.pop();
                    let s = if let Value::Object(obj) = &arr {
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.iter().filter_map(|v| char::from_u32(v.as_i32() as u32)).collect::<String>()
                        } else { String::new() }
                    } else { String::new() };
                    self.push(Value::String(Rc::from(s.as_str())))?;
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
                        Value::Object(o) => {
                            let ob = o.borrow();
                            match &ob.kind {
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                                ObjectKind::Array(_) => "array",
                                _ => "object",
                            }
                        }
                    };
                    self.push(Value::String(Rc::from(tag)))?;
                }
                Op::ref_is_array => {
                    let v = self.pop();
                    let is_arr = matches!(&v, Value::Object(o) if matches!(o.borrow().kind, ObjectKind::Array(_)));
                    self.push(Value::Bool(is_arr))?;
                }

                // -- Array builtins --
                Op::array_length => {
                    let arr = self.pop();
                    let len = if let Value::Object(obj) = &arr {
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind { a.len() as i32 } else { 0 }
                    } else if let Value::String(s) = &arr {
                        s.chars().count() as i32
                    } else { 0 };
                    self.push(Value::I32(len))?;
                }
                Op::array_push => {
                    let val = self.pop(); let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.borrow_mut();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.push(val); }
                    }
                    self.push(arr)?;
                }
                Op::array_pop => {
                    let arr = self.pop();
                    let val = if let Value::Object(obj) = &arr {
                        let mut o = obj.borrow_mut();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.pop().unwrap_or(Value::Null) }
                        else { Value::Null }
                    } else { Value::Null };
                    self.push(val)?;
                }
                Op::array_slice => {
                    let end = self.pop().as_i32(); let start = self.pop().as_i32();
                    let arr = self.pop();
                    let result = if let Value::Object(obj) = &arr {
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind {
                            let len = a.len() as i32;
                            let s = if start < 0 { (len + start).max(0) as usize } else { start.min(len) as usize };
                            let e = if end < 0 { (len + end).max(0) as usize } else { end.min(len) as usize };
                            let sliced: Vec<Value> = a[s..e.max(s)].to_vec();
                            Value::Object(Rc::new(RefCell::new(Object::new_array(sliced))))
                        } else { Value::Null }
                    } else { Value::Null };
                    self.push(result)?;
                }
                Op::array_join => {
                    let delim = self.pop(); let arr = self.pop();
                    let r = if let (Value::Object(obj), Value::String(d)) = (&arr, &delim) {
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind {
                            let parts: Vec<String> = a.iter().map(|v| format!("{}", v)).collect();
                            Rc::from(parts.join(d.as_ref()).as_str())
                        } else { Rc::from("") }
                    } else { Rc::from("") };
                    self.push(Value::String(r))?;
                }
                Op::array_reverse => {
                    let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.borrow_mut();
                        if let ObjectKind::Array(ref mut a) = o.kind { a.reverse(); }
                    }
                    self.push(arr)?;
                }
                Op::array_contains => {
                    let needle = self.pop(); let arr = self.pop();
                    let found = if let Value::Object(obj) = &arr {
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.iter().any(|v| v.eq(&needle))
                        } else { false }
                    } else { false };
                    self.push(Value::Bool(found))?;
                }
                Op::array_index_of => {
                    let needle = self.pop(); let arr = self.pop();
                    let idx = if let Value::Object(obj) = &arr {
                        let o = obj.borrow();
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
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(elems)))))?;
                }
                Op::array_fill => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let start = self.pop().as_i32().max(0) as usize;
                    let val = self.pop();
                    let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.borrow_mut();
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
                        let o = obj.borrow();
                        if let ObjectKind::Array(a) = &o.kind {
                            let end = (src_off + len).min(a.len());
                            a[src_off.min(a.len())..end].to_vec()
                        } else { vec![] }
                    } else { vec![] };
                    // Write to destination
                    if let Value::Object(obj) = &dst {
                        let mut o = obj.borrow_mut();
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
                        let o = obj.borrow();
                        if let ObjectKind::Array(arr) = &o.kind { result.extend(arr.iter().cloned()); }
                    }
                    if let Value::Object(obj) = &b {
                        let o = obj.borrow();
                        if let ObjectKind::Array(arr) = &o.kind { result.extend(arr.iter().cloned()); }
                    }
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(result)))))?;
                }
                Op::array_shift => {
                    let arr = self.pop();
                    let val = if let Value::Object(obj) = &arr {
                        let mut o = obj.borrow_mut();
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
                    obj.properties.insert("__cont_state".into(), Value::String(Rc::from("ready")));
                    obj.properties.insert("__cont_value".into(), Value::Null);
                    self.push(Value::Object(Rc::new(RefCell::new(obj))))?;
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
                            let o = obj.borrow();
                            o.properties.get("__cont_func").cloned().unwrap_or(Value::Null)
                        };
                        {
                            let mut o = obj.borrow_mut();
                            o.properties.insert("__cont_state".into(), Value::String(Rc::from("running")));
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
                        let mut o = obj.borrow_mut();
                        o.properties.insert("__cont_value".into(), val.clone());
                        o.properties.insert("__cont_state".into(), Value::String(Rc::from("running")));
                    }
                    self.push(val)?;
                }

                // -- SIMD (128-bit vectors) --
                Op::v128_load => {
                    let addr = self.pop().as_i32() as usize;
                    let mut bytes = [0u8; 16];
                    if addr + 16 <= self.memory.len() {
                        bytes.copy_from_slice(&self.memory[addr..addr+16]);
                    }
                    self.push(Value::V128(bytes))?;
                }
                Op::v128_store => {
                    let val = self.pop();
                    let addr = self.pop().as_i32() as usize;
                    if let Value::V128(bytes) = val {
                        if addr + 16 <= self.memory.len() {
                            self.memory[addr..addr+16].copy_from_slice(&bytes);
                        }
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
                Op::i32_atomic_load => { let addr = self.pop().as_i32() as usize; if addr+4 <= self.memory.len() { let v = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.push(Value::I32(v))?; } else { self.push(Value::I32(0))?; } }
                // Atomic store: stack is [addr, value] — pop value first (top), then addr
                Op::i32_atomic_store => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { self.memory[addr..addr+4].copy_from_slice(&v.to_le_bytes()); } }
                // Atomic RMW: stack is [addr, value] — pop value (top), addr (second), return old
                Op::i32_atomic_rmw_add => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&old.wrapping_add(v).to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i32_atomic_rmw_sub => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&old.wrapping_sub(v).to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i32_atomic_rmw_and => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&(old & v).to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i32_atomic_rmw_or  => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&(old | v).to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i32_atomic_rmw_xor => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&(old ^ v).to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i32_atomic_rmw_xchg => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); self.memory[addr..addr+4].copy_from_slice(&v.to_le_bytes()); self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                // Atomic CmpXchg: stack is [addr, expected, replacement]
                Op::i32_atomic_rmw_cmpxchg => { let replacement = self.pop().as_i32(); let expected = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; if addr + 4 <= self.memory.len() { let old = i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap()); if old == expected { self.memory[addr..addr+4].copy_from_slice(&replacement.to_le_bytes()); } self.push(Value::I32(old))?; } else { self.push(Value::I32(0))?; } }
                Op::i64_atomic_load => { let addr = self.pop().as_i32() as usize; if addr + 8 <= self.memory.len() { let v = i64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap()); self.push(Value::I64(v))?; } else { self.push(Value::I64(0))?; } }
                Op::i64_atomic_store => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; if addr + 8 <= self.memory.len() { self.memory[addr..addr+8].copy_from_slice(&v.to_le_bytes()); } }
                Op::i64_atomic_rmw_add => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; if addr + 8 <= self.memory.len() { let old = i64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap()); self.memory[addr..addr+8].copy_from_slice(&old.wrapping_add(v).to_le_bytes()); self.push(Value::I64(old))?; } else { self.push(Value::I64(0))?; } }
                Op::i64_atomic_rmw_sub => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; if addr + 8 <= self.memory.len() { let old = i64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap()); self.memory[addr..addr+8].copy_from_slice(&old.wrapping_sub(v).to_le_bytes()); self.push(Value::I64(old))?; } else { self.push(Value::I64(0))?; } }
                Op::i64_atomic_rmw_cmpxchg => { let repl = self.pop().as_i64(); let exp = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; if addr + 8 <= self.memory.len() { let old = i64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap()); if old == exp { self.memory[addr..addr+8].copy_from_slice(&repl.to_le_bytes()); } self.push(Value::I64(old))?; } else { self.push(Value::I64(0))?; } }
                Op::memory_atomic_wait32 => { self.pop(); self.pop(); self.pop(); self.push(Value::I32(1))?; } // not-equal (single-threaded)
                Op::memory_atomic_notify => { self.pop(); self.pop(); self.push(Value::I32(0))?; } // 0 woken (single-threaded)

                // -- Memory64 --
                Op::i64_memory_size => { self.push(Value::I64((self.memory.len() / 65536) as i64))?; }
                Op::i64_memory_grow => { let pages = self.pop().as_i64() as usize; let old = self.memory.len() / 65536; self.memory.resize(self.memory.len() + pages * 65536, 0); self.push(Value::I64(old as i64))?; }
                Op::i32_load_64 => { let addr = self.pop().as_i64() as usize; if addr+4 <= self.memory.len() { self.push(Value::I32(i32::from_le_bytes(self.memory[addr..addr+4].try_into().unwrap())))?; } else { self.push(Value::I32(0))?; } }
                Op::i64_load_64 => { let addr = self.pop().as_i64() as usize; if addr+8 <= self.memory.len() { self.push(Value::I64(i64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap())))?; } else { self.push(Value::I64(0))?; } }
                Op::f64_load_64 => { let addr = self.pop().as_i64() as usize; if addr+8 <= self.memory.len() { self.push(Value::F64(f64::from_le_bytes(self.memory[addr..addr+8].try_into().unwrap())))?; } else { self.push(Value::F64(0.0))?; } }
                Op::i32_store_64 => { let v = self.pop().as_i32(); let addr = self.pop().as_i64() as usize; if addr+4 <= self.memory.len() { self.memory[addr..addr+4].copy_from_slice(&v.to_le_bytes()); } }
                Op::i64_store_64 => { let v = self.pop().as_i64(); let addr = self.pop().as_i64() as usize; if addr+8 <= self.memory.len() { self.memory[addr..addr+8].copy_from_slice(&v.to_le_bytes()); } }
                Op::f64_store_64 => { let v = self.pop().as_f64(); let addr = self.pop().as_i64() as usize; if addr+8 <= self.memory.len() { self.memory[addr..addr+8].copy_from_slice(&v.to_le_bytes()); } }

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
                        let o = obj.borrow();
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

                // -- Stubs --
                Op::iter_get | Op::iter_next | Op::spread => {
                    return Err(VMError::new("Iteration not yet implemented"));
                }
                Op::class_new => { let _ = self.read_u16(); self.push(Value::Null)?; }
                Op::method_def => { let _ = self.read_u16(); }
                Op::inherit => { self.pop(); }
            }
        }
    }

    fn call_value(&mut self, argc: usize) -> Result<(), VMError> {
        let callee_idx = self.stack.len() - 1 - argc;
        let callee = self.stack[callee_idx].clone();

        match &callee {
            Value::Object(obj) => {
                let o = obj.borrow();
                match &o.kind {
                    ObjectKind::Function(func) => {
                        self.call_function(func, argc)?;
                    }
                    ObjectKind::HostFunction(idx) => {
                        // Call host function directly — same as call_import
                        let idx = *idx;
                        drop(o);
                        let args: Vec<Value> = self.stack[self.stack.len() - argc..].to_vec();
                        // Pop args + callee
                        for _ in 0..argc { self.stack.pop(); }
                        self.stack.pop(); // callee
                        let result = (self.host_fns[idx])(&args);
                        self.push(result)?;
                    }
                    _ => return Err(VMError::new("Not a function")),
                }
            }
            _ => return Err(VMError::new(format!("{} is not callable (type: {})", callee.type_tag(), callee))),
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

    fn capture_upvalue(&mut self, stack_idx: usize) -> Rc<RefCell<Upvalue>> {
        for uv in &self.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.borrow().location {
                if idx == stack_idx { return uv.clone(); }
            }
        }
        let uv = Rc::new(RefCell::new(Upvalue { location: UpvalueLocation::Open(stack_idx) }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    fn close_upvalues(&mut self, from: usize) {
        let mut i = 0;
        while i < self.open_upvalues.len() {
            let should_close = matches!(
                self.open_upvalues[i].borrow().location,
                UpvalueLocation::Open(idx) if idx >= from
            );
            if should_close {
                let uv = self.open_upvalues.remove(i);
                let mut u = uv.borrow_mut();
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
                let ob = o.borrow();

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
                let ob = o.borrow();
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
                    Value::Object(Rc::new(RefCell::new(obj)))
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
                let obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0 };
                Value::Object(Rc::new(RefCell::new(obj)))
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
        Value::Object(_) => true,
        Value::V128(b) => b.iter().any(|&x| x != 0),
    }
}
