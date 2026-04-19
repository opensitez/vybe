use std::collections::HashMap;
use std::fmt;
use std::sync::{Arc, Mutex, Weak};

/// A universal VM value. Language-agnostic — no coercion rules here.
/// The compiler is responsible for emitting type checks and conversions.
#[derive(Clone, Debug)]
pub enum Value {
    /// Explicitly empty — VB Nothing, JS null.
    Null,
    /// Uninitialized / missing — JS undefined. VB compiler doesn't emit this.
    Undefined,
    Bool(bool),
    I32(i32),
    I64(i64),
    F64(f64),
    String(Arc<str>),
    Object(Arc<Mutex<Object>>),
    /// Weak reference to an object — does not prevent collection.
    /// Upgrade to strong reference with `ref_deref` opcode.
    WeakRef(Weak<Mutex<Object>>),
    /// SIMD 128-bit vector (4×i32, 2×f64, 4×f32, 16×i8, 8×i16).
    V128([u8; 16]),
    /// JS Symbol — unique identity; description is for debugging only.
    /// Two symbols are `==` only when cloned from the same `Arc<str>`.
    Symbol(Arc<str>),
    /// JS BigInt — arbitrary precision in theory, i64-range in our VM.
    BigInt(i64),
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
            Value::String(s) => s.trim().parse::<f64>().unwrap_or(f64::NAN),
            Value::Null => 0.0,
            _ => f64::NAN,
        }
    }

    pub fn as_i32(&self) -> i32 {
        match self {
            Value::I32(n) => *n,
            Value::I64(n) => *n as i32,
            Value::F64(n) => *n as i32,
            Value::Bool(b) => if *b { 1 } else { 0 },
            Value::String(s) => s.trim().parse::<f64>().map(|f| f as i32).unwrap_or(0),
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
            Value::Symbol(s) => s,
            _ => "",
        }
    }

    /// Value type tag — for the host to inspect.
    pub fn type_tag(&self) -> &'static str {
        match self {
            Value::Null => "null",
            Value::Undefined => "undefined",
            Value::Bool(_) => "bool",
            Value::I32(_) => "i32",
            Value::I64(_) => "i64",
            Value::F64(_) => "f64",
            Value::String(_) => "string",
            Value::Object(o) => {
                let obj = o.lock().unwrap();
                match &obj.kind {
                    ObjectKind::Ordinary => "object",
                    ObjectKind::Array(_) => "array",
                    ObjectKind::Function(_) => "function",
                    ObjectKind::HostFunction(_) => "function",
                }
            }
            Value::V128(_) => "v128",
            Value::WeakRef(_) => "weakref",
            Value::Symbol(_) => "symbol",
            Value::BigInt(_) => "bigint",
        }
    }

    /// Unwrap `Value::BigInt(n)` or coerce narrow integers — used by VM
    /// arithmetic opcodes that route through BigInt.
    pub fn as_bigint(&self) -> i64 {
        match self {
            Value::BigInt(n) => *n,
            Value::I64(n)    => *n,
            Value::I32(n)    => *n as i64,
            _ => 0,
        }
    }

    /// Same-type structural equality. Returns false for different types.
    /// Language-specific equality (JS loose ==, VB implicit conversion)
    /// must be compiled as host calls or bytecode sequences.
    pub fn eq(&self, other: &Value) -> bool {
        match (self, other) {
            (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => true,
            // null == undefined is true in JS loose equality, but this is strict eq
            // JS loose eq is handled by js_loose_eq in the JS compiler
            (Value::Bool(a), Value::Bool(b)) => a == b,
            (Value::I32(a), Value::I32(b)) => a == b,
            (Value::I64(a), Value::I64(b)) => a == b,
            (Value::F64(a), Value::F64(b)) => {
                if a.is_nan() || b.is_nan() { false } else { a == b }
            }
            (Value::String(a), Value::String(b)) => a == b,
            (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
            // Symbols have IDENTITY equality — same Arc instance only.
            (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
            (Value::BigInt(a), Value::BigInt(b)) => a == b,
            (Value::BigInt(a), Value::I64(b)) | (Value::I64(b), Value::BigInt(a)) => *a == *b,
            (Value::BigInt(a), Value::I32(b)) | (Value::I32(b), Value::BigInt(a)) => *a == (*b as i64),
            // Cross-type numeric equality: I32(0) == F64(0.0), etc.
            (Value::I32(a), Value::F64(b)) => (*a as f64) == *b,
            (Value::F64(a), Value::I32(b)) => *a == (*b as f64),
            (Value::I32(a), Value::I64(b)) => (*a as i64) == *b,
            (Value::I64(a), Value::I32(b)) => *a == (*b as i64),
            (Value::I64(a), Value::F64(b)) => (*a as f64) == *b,
            (Value::F64(a), Value::I64(b)) => *a == (*b as f64),
            _ => false,
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Null => write!(f, "null"),
            Value::Undefined => write!(f, "undefined"),
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
                let obj = o.lock().unwrap();
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
            Value::WeakRef(weak) => {
                if weak.upgrade().is_some() {
                    write!(f, "[weakref (alive)]")
                } else {
                    write!(f, "[weakref (dead)]")
                }
            }
            Value::V128(bytes) => {
                let vals: Vec<String> = bytes.iter().map(|b| format!("{:02x}", b)).collect();
                write!(f, "v128[{}]", vals.join(""))
            }
            Value::Symbol(d) => write!(f, "Symbol({})", d),
            Value::BigInt(n)  => write!(f, "{}n", n),
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
    /// Indexed fields — fixed-layout storage for typed objects (WASM GC struct fields).
    /// Field i is accessed by index when the type's field layout is known.
    /// Dynamic properties spill into `properties` HashMap.
    pub fields: Vec<Value>,
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
        Object { properties: HashMap::new(), kind: ObjectKind::Ordinary, type_id: 0, fields: Vec::new() }
    }

    pub fn new_typed(type_id: usize) -> Self {
        Object { properties: HashMap::new(), kind: ObjectKind::Ordinary, type_id, fields: Vec::new() }
    }

    /// Create a typed object with pre-allocated indexed fields.
    pub fn new_typed_with_fields(type_id: usize, field_count: usize) -> Self {
        Object {
            properties: HashMap::new(),
            kind: ObjectKind::Ordinary,
            type_id,
            fields: vec![Value::Null; field_count],
        }
    }

    pub fn new_array(elements: Vec<Value>) -> Self {
        let len = elements.len();
        let mut obj = Object {
            properties: HashMap::new(),
            kind: ObjectKind::Array(elements),
            type_id: 0,
            fields: Vec::new(),
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
    pub upvalues: Vec<Arc<Mutex<Upvalue>>>,
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
