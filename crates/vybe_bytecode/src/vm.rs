use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::chunk::Chunk;
use crate::error::VMError;
use crate::opcode::Op;
use crate::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation, Value};

const MAX_FRAMES: usize = 256;
const MAX_STACK: usize = 65536;

/// Host function signature. Receives args, returns a value.
pub type HostFn = Box<dyn Fn(&[Value]) -> Value>;

#[derive(Debug, Clone)]
struct CallFrame {
    chunk_index: usize,
    ip: usize,
    base: usize,
    upvalues: Vec<Rc<RefCell<Upvalue>>>,
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
    host_fn_names: Vec<String>,
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
            host_fn_names: Vec::new(),
        }
    }

    /// Register a host function. Returns its index (used by CallHost opcode).
    pub fn register_host_fn(&mut self, name: impl Into<String>, f: HostFn) -> u16 {
        let idx = self.host_fns.len() as u16;
        self.host_fn_names.push(name.into());
        self.host_fns.push(f);
        idx
    }

    /// Load chunks and execute chunk 0 (the script).
    pub fn run(&mut self, chunks: Vec<Chunk>) -> Result<Value, VMError> {
        self.chunks = chunks;
        if self.chunks.is_empty() {
            return Ok(Value::Null);
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

        self.execute()
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
                Op::Halt => {
                    // Close all open upvalues so closures retain captured values
                    self.close_upvalues(0);
                    return Ok(if self.stack.is_empty() { Value::Null } else { self.pop() });
                }

                Op::Const => {
                    let idx = self.read_u16();
                    let val = self.get_constant(idx);
                    self.push(val)?;
                }
                Op::Pop => { self.pop(); }
                Op::Dup => {
                    let val = self.peek(0).clone();
                    self.push(val)?;
                }

                // -- Variables --
                Op::GetLocal => {
                    let slot = self.read_u16() as usize;
                    let base = self.frame().base;
                    let val = self.stack[base + slot].clone();
                    self.push(val)?;
                }
                Op::SetLocal => {
                    let slot = self.read_u16() as usize;
                    let val = self.peek(0).clone();
                    let base = self.frame().base;
                    self.stack[base + slot] = val;
                }
                Op::GetGlobal => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.globals.get(&name).cloned().unwrap_or(Value::Null);
                    self.push(val)?;
                }
                Op::SetGlobal => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.peek(0).clone();
                    self.globals.insert(name, val);
                }
                Op::GetUpvalue => {
                    let idx = self.read_byte() as usize;
                    let uv = self.frame().upvalues[idx].clone();
                    let val = match &uv.borrow().location {
                        UpvalueLocation::Open(si) => self.stack[*si].clone(),
                        UpvalueLocation::Closed(v) => v.clone(),
                    };
                    self.push(val)?;
                }
                Op::SetUpvalue => {
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
                Op::GetProp => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let obj = self.pop();
                    match &obj {
                        Value::Object(o) => {
                            let val = o.borrow().get(&name);
                            self.push(val)?;
                        }
                        Value::String(s) if name == "length" => {
                            self.push(Value::F64(s.len() as f64))?;
                        }
                        _ => self.push(Value::Null)?,
                    }
                }
                Op::SetProp => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.pop();
                    let obj = self.pop();
                    if let Value::Object(o) = &obj {
                        o.borrow_mut().set(name, val.clone());
                    }
                    self.push(val)?;
                }
                Op::GetIndex => {
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
                Op::SetIndex => {
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
                Op::AddF => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a + b))?;
                }
                Op::SubF => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a - b))?;
                }
                Op::MulF => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a * b))?;
                }
                Op::DivF => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a / b))?;
                }
                Op::ModF => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a % b))?;
                }
                Op::NegF => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(-a))?;
                }

                // -- Integer arithmetic --
                Op::AddI => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_add(b)))?;
                }
                Op::SubI => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_sub(b)))?;
                }
                Op::MulI => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_mul(b)))?;
                }

                // -- String --
                Op::Concat => {
                    let b = self.pop();
                    let a = self.pop();
                    let s = format!("{}{}", a, b);
                    self.push(Value::String(Rc::from(s.as_str())))?;
                }
                Op::StrConcat => {
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
                Op::BitAnd => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a & b))?; }
                Op::BitOr  => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a | b))?; }
                Op::BitXor => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a ^ b))?; }
                Op::BitNot => { let a = self.pop().as_i32(); self.push(Value::I32(!a))?; }
                Op::Shl    => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a << (b & 0x1f)))?; }
                Op::Shr    => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a >> (b & 0x1f)))?; }
                Op::UShr   => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32((a >> (b & 0x1f)) as i32))?; }

                // -- Comparison --
                Op::CmpEq => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.eq(&b)))?;
                }
                Op::CmpNe => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(!a.eq(&b)))?;
                }
                Op::CmpLtF => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a < b))?; }
                Op::CmpGtF => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a > b))?; }
                Op::CmpLeF => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a <= b))?; }
                Op::CmpGeF => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a >= b))?; }
                Op::CmpLtS => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.as_str() < b.as_str()))?;
                }
                Op::CmpGtS => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.as_str() > b.as_str()))?;
                }

                // -- Logical --
                Op::BoolNot => {
                    let a = self.pop().as_bool();
                    self.push(Value::Bool(!a))?;
                }

                // -- Control flow --
                Op::Jump => {
                    let offset = self.read_i16();
                    let f = self.frame_mut();
                    f.ip = (f.ip as i64 + offset as i64) as usize;
                }
                Op::JumpIfFalse => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == false {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                Op::JumpIfTrue => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == true {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                Op::JumpIfNull => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if matches!(val, Value::Null) {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }

                // -- Functions --
                Op::Call => {
                    let argc = self.read_byte() as usize;
                    self.call_value(argc)?;
                }
                Op::Return => {
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
                Op::Closure => {
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
                    let obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func) };
                    self.push(Value::Object(Rc::new(RefCell::new(obj))))?;
                }

                // -- Host functions --
                Op::CallHost => {
                    let fn_idx = self.read_u16() as usize;
                    let argc = self.read_byte() as usize;
                    let base = self.stack.len() - argc;
                    let args: Vec<Value> = self.stack[base..].to_vec();
                    self.stack.truncate(base);

                    if fn_idx < self.host_fns.len() {
                        let result = (self.host_fns[fn_idx])(&args);
                        self.push(result)?;
                    } else {
                        return Err(VMError::new(format!("Unknown host function index: {}", fn_idx)));
                    }
                }

                // -- Object/Array --
                Op::NewObject => {
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
                Op::NewArray => {
                    let count = self.read_u16() as usize;
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    self.push(Value::Object(Rc::new(RefCell::new(Object::new_array(elems)))))?;
                }

                // -- Immediates --
                Op::PushNull => self.push(Value::Null)?,
                Op::PushTrue => self.push(Value::Bool(true))?,
                Op::PushFalse => self.push(Value::Bool(false))?,
                Op::PushI32Zero => self.push(Value::I32(0))?,
                Op::PushI32One => self.push(Value::I32(1))?,
                Op::PushF64Zero => self.push(Value::F64(0.0))?,

                // -- Type checks --
                Op::IsNull => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Null)))?; }
                Op::IsString => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::String(_))))?; }
                Op::IsNumber => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::F64(_) | Value::I32(_) | Value::I64(_))))?; }
                Op::IsBool => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Bool(_))))?; }
                Op::IsObject => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Object(_))))?; }
                Op::IsFunction => {
                    let v = self.pop();
                    let is_fn = matches!(&v, Value::Object(o) if matches!(o.borrow().kind, ObjectKind::Function(_)));
                    self.push(Value::Bool(is_fn))?;
                }

                // -- Conversions --
                Op::ToF64 => {
                    let v = self.pop();
                    self.push(Value::F64(v.as_f64()))?;
                }
                Op::ToI32 => {
                    let v = self.pop();
                    self.push(Value::I32(v.as_i32()))?;
                }

                // -- Exceptions --
                Op::TryStart => { let _ = self.read_u16(); let _ = self.read_u16(); }
                Op::TryEnd => {}
                Op::Throw => {
                    let val = self.pop();
                    return Err(VMError::new(format!("{}", val)));
                }

                // -- Stubs --
                Op::GetIterator | Op::IterNext | Op::Spread => {
                    return Err(VMError::new("Iteration not yet implemented"));
                }
                Op::Class => { let _ = self.read_u16(); self.push(Value::Null)?; }
                Op::Method => { let _ = self.read_u16(); }
                Op::Inherit => { self.pop(); }
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
                    _ => return Err(VMError::new("Not a function")),
                }
            }
            _ => return Err(VMError::new(format!("{} is not callable", callee.type_tag()))),
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
