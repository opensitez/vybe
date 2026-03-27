use std::cell::RefCell;
use std::collections::HashMap;
use std::fmt;
use std::rc::Rc;

/// A universal VM value. Language-agnostic — no coercion rules here.
/// The compiler is responsible for emitting type checks and conversions.
#[derive(Clone, Debug)]
pub enum Value {
    /// No value / null / nothing / nil — universal "absence" marker.
    Null,
    Bool(bool),
    I32(i32),
    I64(i64),
    F64(f64),
    String(Rc<str>),
    Object(Rc<RefCell<Object>>),
}

impl Value {
    /// Extract f64 or panic. VM arithmetic ops require the compiler
    /// to have already ensured the operand is numeric.
    pub fn as_f64(&self) -> f64 {
        match self {
            Value::F64(n) => *n,
            Value::I32(n) => *n as f64,
            Value::I64(n) => *n as f64,
            Value::Bool(b) => if *b { 1.0 } else { 0.0 },
            _ => f64::NAN,
        }
    }

    pub fn as_i32(&self) -> i32 {
        match self {
            Value::I32(n) => *n,
            Value::I64(n) => *n as i32,
            Value::F64(n) => *n as i32,
            Value::Bool(b) => if *b { 1 } else { 0 },
            _ => 0,
        }
    }

    pub fn as_i64(&self) -> i64 {
        match self {
            Value::I64(n) => *n,
            Value::I32(n) => *n as i64,
            Value::F64(n) => *n as i64,
            Value::Bool(b) => if *b { 1 } else { 0 },
            _ => 0,
        }
    }

    pub fn as_bool(&self) -> bool {
        matches!(self, Value::Bool(true))
    }

    pub fn as_str(&self) -> &str {
        match self {
            Value::String(s) => s,
            _ => "",
        }
    }

    /// Value type tag — for the host to inspect.
    pub fn type_tag(&self) -> &'static str {
        match self {
            Value::Null => "null",
            Value::Bool(_) => "bool",
            Value::I32(_) => "i32",
            Value::I64(_) => "i64",
            Value::F64(_) => "f64",
            Value::String(_) => "string",
            Value::Object(o) => {
                let obj = o.borrow();
                match &obj.kind {
                    ObjectKind::Ordinary => "object",
                    ObjectKind::Array(_) => "array",
                    ObjectKind::Function(_) => "function",
                    ObjectKind::HostFunction(_) => "function",
                }
            }
        }
    }

    /// Same-type structural equality. Returns false for different types.
    /// Language-specific equality (JS loose ==, VB implicit conversion)
    /// must be compiled as host calls or bytecode sequences.
    pub fn eq(&self, other: &Value) -> bool {
        match (self, other) {
            (Value::Null, Value::Null) => true,
            (Value::Bool(a), Value::Bool(b)) => a == b,
            (Value::I32(a), Value::I32(b)) => a == b,
            (Value::I64(a), Value::I64(b)) => a == b,
            (Value::F64(a), Value::F64(b)) => {
                if a.is_nan() || b.is_nan() { false } else { a == b }
            }
            (Value::String(a), Value::String(b)) => a == b,
            (Value::Object(a), Value::Object(b)) => Rc::ptr_eq(a, b),
            _ => false,
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Null => write!(f, "null"),
            Value::Bool(b) => write!(f, "{}", b),
            Value::I32(n) => write!(f, "{}", n),
            Value::I64(n) => write!(f, "{}", n),
            Value::F64(n) => {
                if *n == (*n as i64) as f64 && n.abs() < 1e15 && !n.is_infinite() && !n.is_nan() {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{}", n)
                }
            }
            Value::String(s) => write!(f, "{}", s),
            Value::Object(o) => {
                let obj = o.borrow();
                match &obj.kind {
                    ObjectKind::Array(elems) => {
                        let parts: Vec<String> = elems.iter().map(|v| format!("{}", v)).collect();
                        write!(f, "{}", parts.join(","))
                    }
                    ObjectKind::Function(func) => {
                        write!(f, "[function {}]", func.name.as_deref().unwrap_or("anonymous"))
                    }
                    ObjectKind::HostFunction(idx) => write!(f, "[host function {}]", idx),
                    ObjectKind::Ordinary => write!(f, "[object]"),
                }
            }
        }
    }
}

/// A heap-allocated object with named properties and an internal kind.
#[derive(Debug, Clone)]
pub struct Object {
    pub properties: HashMap<String, Value>,
    pub kind: ObjectKind,
    /// WASM GC-style type reference. 0 = Object (untyped), >0 = specific type.
    pub type_id: usize,
}

#[derive(Debug, Clone)]
pub enum ObjectKind {
    Ordinary,
    Array(Vec<Value>),
    Function(Function),
    /// A reference to a host function by its index in the VM's host_fns table.
    HostFunction(usize),
}

impl Object {
    pub fn new() -> Self {
        Object { properties: HashMap::new(), kind: ObjectKind::Ordinary, type_id: 0 }
    }

    pub fn new_typed(type_id: usize) -> Self {
        Object { properties: HashMap::new(), kind: ObjectKind::Ordinary, type_id }
    }

    pub fn new_array(elements: Vec<Value>) -> Self {
        let len = elements.len();
        let mut obj = Object {
            properties: HashMap::new(),
            kind: ObjectKind::Array(elements),
            type_id: 0, // Will be set to Array/List type_id by the host
        };
        obj.properties.insert("length".into(), Value::F64(len as f64));
        obj
    }

    pub fn get(&self, key: &str) -> Value {
        if let Some(v) = self.properties.get(key) {
            return v.clone();
        }
        if let ObjectKind::Array(ref elems) = self.kind {
            if let Ok(idx) = key.parse::<usize>() {
                if idx < elems.len() {
                    return elems[idx].clone();
                }
            }
        }
        Value::Null
    }

    pub fn set(&mut self, key: String, value: Value) {
        if let ObjectKind::Array(ref mut elems) = self.kind {
            if let Ok(idx) = key.parse::<usize>() {
                if idx >= elems.len() {
                    elems.resize(idx + 1, Value::Null);
                }
                elems[idx] = value.clone();
                self.properties.insert("length".into(), Value::F64(elems.len() as f64));
                return;
            }
        }
        self.properties.insert(key, value);
    }
}

/// A bytecode function (closure).
#[derive(Debug, Clone)]
pub struct Function {
    pub name: Option<String>,
    pub arity: u8,
    pub chunk_index: usize,
    pub upvalues: Vec<Rc<RefCell<Upvalue>>>,
}

/// A captured variable (upvalue).
#[derive(Debug, Clone)]
pub struct Upvalue {
    pub location: UpvalueLocation,
}

#[derive(Debug, Clone)]
pub enum UpvalueLocation {
    Open(usize),
    Closed(Value),
}
