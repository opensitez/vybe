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


    /// Load chunks and execute chunk 0 (the script).
    /// Resolves the import table against registered host functions.
    pub fn run(&mut self, chunks: Vec<Chunk>) -> Result<Value, VMError> {
        self.chunks = chunks;
        if self.chunks.is_empty() {
            return Ok(Value::Null);
        }

        // Resolve imports from chunk 0's import table
        self.import_table.clear();
        for import in &self.chunks[0].imports {
            let key = (import.module.clone(), import.name.clone());
            match self.host_registry.get(&key) {
                Some(&idx) => self.import_table.push(idx),
                None => return Err(VMError::new(format!(
                    "Unresolved import: \"{}\" \"{}\"", import.module, import.name
                ))),
            }
        }

        self.frames.push(CallFrame {
            chunk_index: 0,
            ip: 0,
            base: 0,
            upvalues: Vec::new(),
        });

        let local_count = self.chunks[0].local_count as usize;
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
            let op = match Op::from_byte(byte) {
                Some(op) => op,
                None => return Err(VMError::new(format!("Invalid opcode: {}", byte))),
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
                    let val = self.globals.get(&name).cloned().unwrap_or(Value::Null);
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

                // -- String --
                Op::str_concat => {
                    let b = self.pop();
                    let a = self.pop();
                    let s = format!("{}{}", a, b);
                    self.push(Value::String(Rc::from(s.as_str())))?;
                }
                Op::str_concat_n => {
                    let count = self.read_byte() as usize;
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
                        self.push(result)?;
                    } else {
                        return Err(VMError::new(format!("Unresolved import index: {}", import_idx)));
                    }
                }

                // -- Object/Array --
                Op::struct_new => {
                    let count = self.read_u16() as usize;
                    let mut obj = Object::new();
                    let start = self.stack.len() - count * 2;
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
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(elems)))))?;
                }

                // -- Immediates --
                Op::null => self.push(Value::Null)?,
                Op::r#true => self.push(Value::Bool(true))?,
                Op::r#false => self.push(Value::Bool(false))?,
                Op::i32_const_0 => self.push(Value::I32(0))?,
                Op::i32_const_1 => self.push(Value::I32(1))?,
                Op::f64_const_0 => self.push(Value::F64(0.0))?,

                // -- Type checks --
                Op::ref_is_null => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Null)))?; }
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
                        (Value::Null, Value::Null) => true,
                        (Value::Bool(x), Value::Bool(y)) => x == y,
                        (Value::F64(x), Value::F64(y)) => if x.is_nan() || y.is_nan() { false } else { x == y },
                        (Value::I32(x), Value::I32(y)) => x == y,
                        (Value::F64(x), Value::I32(y)) => *x == *y as f64,
                        (Value::I32(x), Value::F64(y)) => *x as f64 == *y,
                        (Value::String(x), Value::String(y)) => x == y,
                        (Value::Object(x), Value::Object(y)) => Rc::ptr_eq(x, y),
                        _ => false,
                    };
                    self.push(Value::Bool(result))?;
                }
                Op::dyn_ne => {
                    let b = self.pop(); let a = self.pop();
                    self.push(Value::Bool(!a.eq(&b)))?;
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
                    self.push(Value::Bool(a.as_f64() <= b.as_f64()))?;
                }
                Op::dyn_ge => {
                    let b = self.pop(); let a = self.pop();
                    self.push(Value::Bool(a.as_f64() >= b.as_f64()))?;
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
        Value::Null => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        Value::Object(_) => true,
    }
}
