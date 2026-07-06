use std::collections::HashMap;
use std::fmt;
use std::hash::{Hash, Hasher};
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
    /// JS BigInt host value — ARBITRARY PRECISION per ECMA-262 §6.1.6.2
    /// (js-primitive-builtins models bigint as an opaque host type; the
    /// js-types JS-API converts wasm i64 ⇄ BigInt via ToBigInt64, the
    /// only place a 64-bit wrap is legal). Arc: clone = refcount bump.
    BigInt(crate::bigint::BigIntRef),
}

/// Compact tag identifying the `Value` variant — a small integer that
/// indexes into per-slot counter arrays in the VM's type recorder.
/// Keep in sync with the variants above; new variants append to the
/// end so existing counter indices stay stable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum ValueTag {
    Null = 0,
    Undefined = 1,
    Bool = 2,
    I32 = 3,
    I64 = 4,
    F64 = 5,
    String = 6,
    Object = 7,
    WeakRef = 8,
    V128 = 9,
    Symbol = 10,
    BigInt = 11,
}

impl ValueTag {
    pub const COUNT: usize = 12;
    pub fn as_usize(self) -> usize {
        self as usize
    }
    pub fn name(self) -> &'static str {
        match self {
            ValueTag::Null => "Null",
            ValueTag::Undefined => "Undefined",
            ValueTag::Bool => "Bool",
            ValueTag::I32 => "I32",
            ValueTag::I64 => "I64",
            ValueTag::F64 => "F64",
            ValueTag::String => "String",
            ValueTag::Object => "Object",
            ValueTag::WeakRef => "WeakRef",
            ValueTag::V128 => "V128",
            ValueTag::Symbol => "Symbol",
            ValueTag::BigInt => "BigInt",
        }
    }
}

impl Value {
    /// Compact tag identifying this variant. Used by the type recorder
    /// to index into per-slot counter arrays — avoids a HashMap-per-slot
    /// and keeps recording cheap enough to leave on during test runs.
    pub fn tag(&self) -> ValueTag {
        match self {
            Value::Null => ValueTag::Null,
            Value::Undefined => ValueTag::Undefined,
            Value::Bool(_) => ValueTag::Bool,
            Value::I32(_) => ValueTag::I32,
            Value::I64(_) => ValueTag::I64,
            Value::F64(_) => ValueTag::F64,
            Value::String(_) => ValueTag::String,
            Value::Object(_) => ValueTag::Object,
            Value::WeakRef(_) => ValueTag::WeakRef,
            Value::V128(_) => ValueTag::V128,
            Value::Symbol(_) => ValueTag::Symbol,
            Value::BigInt(_) => ValueTag::BigInt,
        }
    }
}

impl Value {
    /// Extract f64 or panic. VM arithmetic ops require the compiler
    /// to have already ensured the operand is numeric.
    pub fn as_f64(&self) -> f64 {
        match self {
            Value::F64(n) => *n,
            Value::I32(n) => *n as f64,
            Value::I64(n) => *n as f64,
            Value::Bool(b) => {
                if *b {
                    1.0
                } else {
                    0.0
                }
            }
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
            Value::Bool(b) => {
                if *b {
                    1
                } else {
                    0
                }
            }
            Value::String(s) => s.trim().parse::<f64>().map(|f| f as i32).unwrap_or(0),
            _ => 0,
        }
    }

    /// ECMA-262 ToInt32 (§7.1.6): truncate to integer, reduce modulo 2^32,
    /// then interpret as signed 32-bit. This wraps instead of saturating,
    /// matching the semantics required by JS bitwise operators.
    pub fn to_ecma_int32(&self) -> i32 {
        let n = self.as_f64();
        if n.is_nan() || n.is_infinite() {
            return 0;
        }
        n.trunc().rem_euclid(4_294_967_296.0) as u64 as u32 as i32
    }

    /// ECMA-262 ToUint32 (§7.1.7): same as ToInt32 but interpret as unsigned.
    pub fn to_ecma_uint32(&self) -> u32 {
        self.to_ecma_int32() as u32
    }

    /// `Value::BigInt` from an i64 (exact — no precision involved).
    pub fn bigint_i64(n: i64) -> Value {
        Value::BigInt(Arc::new(crate::bigint::BigIntVal::from_i64(n)))
    }

    /// `Value::BigInt` from a u64 (exact — ToBigUint64 reading).
    pub fn bigint_u64(n: u64) -> Value {
        Value::BigInt(Arc::new(crate::bigint::BigIntVal::from_u64(n)))
    }

    /// `Value::BigInt` from an owned arbitrary-precision value.
    pub fn bigint(v: crate::bigint::BigIntVal) -> Value {
        Value::BigInt(Arc::new(v))
    }

    pub fn as_i64(&self) -> i64 {
        match self {
            Value::I64(n) => *n,
            // ToBigInt64 semantics: wrap modulo 2^64 (JS-API i64 boundary).
            Value::BigInt(n) => n.to_i64_wrapping(),
            Value::I32(n) => *n as i64,
            Value::F64(n) => *n as i64,
            Value::Bool(b) => {
                if *b {
                    1
                } else {
                    0
                }
            }
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
                    ObjectKind::Map(_) => "map",
                    ObjectKind::Set(_) => "set",
                    ObjectKind::ArrayBuffer(_) => "arraybuffer",
                    ObjectKind::TypedArray(_) => "typedarray",
                    ObjectKind::Function(_) => "function",
                    ObjectKind::HostFunction(_) => "function",
                    ObjectKind::ModuleNamespace => "object",
                    ObjectKind::Continuation(_) => "continuation",
                    ObjectKind::Future { .. } => "future",
                    ObjectKind::Stream { .. } => "stream",
                }
            }
            Value::V128(_) => "v128",
            Value::WeakRef(_) => "weakref",
            Value::Symbol(_) => "symbol",
            Value::BigInt(_) => "bigint",
        }
    }

    /// Unwrap `Value::BigInt(n)` or coerce narrow integers — used by VM
    /// arithmetic opcodes that route through BigInt. 64-bit view:
    /// ToBigInt64 wrap for values beyond i64 (the i64.* ops are 64-bit).
    pub fn as_bigint(&self) -> i64 {
        match self {
            Value::BigInt(n) => n.to_i64_wrapping(),
            Value::I64(n) => *n,
            Value::I32(n) => *n as i64,
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
                if a.is_nan() || b.is_nan() {
                    false
                } else {
                    a == b
                }
            }
            (Value::String(a), Value::String(b)) => a == b,
            (Value::Object(a), Value::Object(b)) => {
                if Arc::ptr_eq(a, b) {
                    true
                } else {
                    let oa = a.lock().unwrap();
                    let ob = b.lock().unwrap();
                    let wrapper_call_a = oa
                        .properties
                        .get("__call__")
                        .cloned()
                        .or_else(|| oa.properties.get("call").cloned());
                    let wrapper_call_b = ob
                        .properties
                        .get("__call__")
                        .cloned()
                        .or_else(|| ob.properties.get("call").cloned());
                    let kind_eq = match (&oa.kind, &ob.kind) {
                        (ObjectKind::Function(fa), ObjectKind::Function(fb)) => {
                            fa.chunk_index == fb.chunk_index
                        }
                        (ObjectKind::HostFunction(ia), ObjectKind::HostFunction(ib)) => ia == ib,
                        _ => false,
                    };
                    drop(oa);
                    drop(ob);

                    if kind_eq {
                        true
                    } else if let (Some(ca), Some(cb)) = (wrapper_call_a, wrapper_call_b) {
                        ca.eq(&cb)
                    } else {
                        false
                    }
                }
            }
            // Symbols have IDENTITY equality — same Arc instance only.
            (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
            (Value::BigInt(a), Value::BigInt(b)) => a == b,
            (Value::BigInt(a), Value::I64(b)) | (Value::I64(b), Value::BigInt(a)) => {
                **a == crate::bigint::BigIntVal::from_i64(*b)
            }
            (Value::BigInt(a), Value::I32(b)) | (Value::I32(b), Value::BigInt(a)) => {
                **a == crate::bigint::BigIntVal::from_i64(*b as i64)
            }
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

// ── SameValueZero-style hashing / equality ──────────────────────────────
//
// `PartialEq + Eq + Hash` on `Value` implement the JS `Map` / `Set` key
// semantics (ECMA-262 §7.2.11 SameValueZero):
//
//   * `NaN` compares equal to itself (only case where this differs
//     from the usual `===` contract).
//   * `-0` and `+0` are considered the same key (no sign distinction
//     in `Map` keys — `Object.is` is the exception that differs).
//   * Numeric values across `I32` / `I64` / `F64` with the same
//     integral value hash and compare the same bucket so that
//     `map.set(1.0, v); map.get(1)` works as a JS programmer expects.
//   * Strings compare by content, `Arc` clones hash to the same bucket.
//   * `Object` / `Symbol` / `WeakRef` compare by pointer identity,
//     matching v8's behavior for object keys.
//
// These impls are REQUIRED for `indexmap::IndexMap<Value, Value>`
// backing the `ObjectKind::Map` / `ObjectKind::Set` variants
// (Phase B4 of `dynamicruntime_support.md`).

impl std::cmp::PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        Value::same_value_zero(self, other)
    }
}

impl std::cmp::Eq for Value {}

impl Value {
    /// ECMA-262 §7.2.11 SameValueZero — the algorithm used for
    /// `Array.prototype.includes`, `Map` key equality, `Set` member
    /// equality. Differs from `===` (strict equality) only in that
    /// `NaN` is equal to itself. Differs from `Object.is` in that
    /// `-0` and `+0` are equal here but distinct under `Object.is`.
    pub fn same_value_zero(a: &Value, b: &Value) -> bool {
        match (a, b) {
            (Value::Null, Value::Null) => true,
            (Value::Undefined, Value::Undefined) => true,
            (Value::Bool(x), Value::Bool(y)) => x == y,

            // Numeric cross-type coercion — a JS user writes
            // `map.set(1, "a")` and expects `map.get(1.0)` to return it.
            (Value::I32(x), Value::I32(y)) => x == y,
            (Value::I64(x), Value::I64(y)) => x == y,
            (Value::I32(x), Value::I64(y)) => (*x as i64) == *y,
            (Value::I64(x), Value::I32(y)) => *x == (*y as i64),
            (Value::F64(x), Value::F64(y)) => {
                if x.is_nan() && y.is_nan() {
                    true
                } else {
                    x == y
                }
            }
            (Value::I32(x), Value::F64(y)) => !y.is_nan() && (*x as f64) == *y,
            (Value::F64(x), Value::I32(y)) => !x.is_nan() && *x == (*y as f64),
            (Value::I64(x), Value::F64(y)) => !y.is_nan() && (*x as f64) == *y,
            (Value::F64(x), Value::I64(y)) => !x.is_nan() && *x == (*y as f64),

            (Value::String(x), Value::String(y)) => x == y,
            (Value::BigInt(x), Value::BigInt(y)) => x == y,
            (Value::V128(x), Value::V128(y)) => x == y,

            // Pointer identity for reference types.
            (Value::Object(x), Value::Object(y)) => Arc::ptr_eq(x, y),
            (Value::Symbol(x), Value::Symbol(y)) => Arc::ptr_eq(x, y),
            (Value::WeakRef(x), Value::WeakRef(y)) => Weak::ptr_eq(x, y),

            _ => false,
        }
    }
}

impl Hash for Value {
    fn hash<H: Hasher>(&self, state: &mut H) {
        // Consistency with `PartialEq::eq` above:
        // any two values that compare equal must hash to the same
        // bucket. Numeric cross-type equality means I32(1), I64(1),
        // and F64(1.0) all hash as the same i64 under one tag.
        match self {
            Value::Null => 0u8.hash(state),
            Value::Undefined => 1u8.hash(state),
            Value::Bool(b) => {
                2u8.hash(state);
                b.hash(state);
            }
            Value::I32(n) => {
                3u8.hash(state);
                (*n as i64).hash(state);
            }
            Value::I64(n) => {
                3u8.hash(state);
                n.hash(state);
            }
            Value::F64(n) => {
                3u8.hash(state);
                if n.is_nan() {
                    // All NaN bit-patterns must hash the same so
                    // SameValueZero's NaN === NaN equality stays
                    // consistent.
                    i64::MIN.hash(state);
                } else if *n == 0.0 {
                    // -0.0 / +0.0 collapse to same bucket.
                    0i64.hash(state);
                } else if n.fract() == 0.0 && *n >= i64::MIN as f64 && *n <= i64::MAX as f64 {
                    // Integral float — hash as the i64 it equals,
                    // so F64(5.0) and I32(5) share a bucket.
                    (*n as i64).hash(state);
                } else {
                    // Non-integral float — use a separate subspace
                    // keyed by the raw bit pattern. A non-integral
                    // f64 can never equal an i32/i64, so no
                    // cross-bucket collision risk.
                    u8::MAX.hash(state);
                    n.to_bits().hash(state);
                }
            }
            Value::String(s) => {
                5u8.hash(state);
                s.hash(state);
            }
            Value::Object(o) => {
                6u8.hash(state);
                (Arc::as_ptr(o) as usize).hash(state);
            }
            Value::WeakRef(w) => {
                // Hash by the weak pointer's raw address. `Weak::as_ptr`
                // stays stable across upgrade/drop transitions.
                7u8.hash(state);
                (Weak::as_ptr(w) as usize).hash(state);
            }
            Value::V128(bytes) => {
                8u8.hash(state);
                bytes.hash(state);
            }
            Value::Symbol(s) => {
                9u8.hash(state);
                // `Arc<str>` is a fat pointer; cast through a thin
                // pointer to get a stable usize identity.
                (Arc::as_ptr(s) as *const u8 as usize).hash(state);
            }
            Value::BigInt(n) => {
                10u8.hash(state);
                n.hash(state);
            }
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
                if n.is_nan() {
                    write!(f, "NaN")
                } else if n.is_infinite() {
                    if n.is_sign_negative() {
                        write!(f, "-Infinity")
                    } else {
                        write!(f, "Infinity")
                    }
                } else if *n == (*n as i64) as f64 && n.abs() < 1e15 {
                    // NOTE: -0.0 deliberately prints "0" — Display IS the
                    // §6.1.6.1.20 ToString surface here (String(-0)==="0").
                    // Node's console shows "-0" via its inspector, which
                    // would need a console-path formatter, never Display.
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
                    ObjectKind::Map(m) => {
                        // ECMA-262 §24.1.3.13: no canonical toString for
                        // Map; v8 shows "[object Map]". Match that.
                        let _ = m;
                        write!(f, "[object Map]")
                    }
                    ObjectKind::Set(s) => {
                        let _ = s;
                        write!(f, "[object Set]")
                    }
                    ObjectKind::ArrayBuffer(ab) => {
                        let _ = ab;
                        write!(f, "[object ArrayBuffer]")
                    }
                    ObjectKind::TypedArray(ta) => {
                        // Spec: typed arrays toString as comma-joined elements.
                        // MVP: tag-only — proper toString lives in the handler.
                        let _ = ta;
                        write!(f, "[object TypedArray]")
                    }
                    ObjectKind::Function(func) => {
                        write!(
                            f,
                            "[function {}]",
                            func.name.as_deref().unwrap_or("anonymous")
                        )
                    }
                    ObjectKind::HostFunction(idx) => write!(f, "[host function {}]", idx),
                    // Per ECMA-262 §10.4.6 the `Symbol.toStringTag`
                    // own property is `"Module"`, so `Object.prototype.toString.call(ns)`
                    // returns `"[object Module]"`.
                    ObjectKind::ModuleNamespace => write!(f, "[object Module]"),
                    ObjectKind::Continuation(_) => write!(f, "[continuation]"),
                    ObjectKind::Future { id } => write!(f, "[future {}]", id),
                    ObjectKind::Stream { id } => write!(f, "[stream {}]", id),
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
            Value::BigInt(n) => write!(f, "{}n", n),
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

/// Backing state for an `ObjectKind::ArrayBuffer`. Carries the bytes
/// plus the resizability / detachment metadata required by
/// ECMA-262 §25.1 (`ArrayBuffer`) and §25.2 (`SharedArrayBuffer`).
///
/// `bytes` is stored as `Arc<Mutex<Vec<u8>>>` so that `DataView` and
/// `TypedArray` views can share the underlying storage — a write
/// through any view is observable via every other view on the same
/// buffer, matching the JS spec contract.
#[derive(Debug, Clone)]
pub struct ArrayBufferState {
    pub bytes: Arc<Mutex<Vec<u8>>>,
    pub max_byte_length: usize,
    pub resizable: bool,
    /// True if `transfer()` has rendered this buffer unusable.
    pub detached: bool,
    /// True for `SharedArrayBuffer`. Affects whether cross-thread
    /// writes are allowed; the MVP implementation doesn't yet enforce
    /// this differently from a non-shared buffer.
    pub shared: bool,
}

/// Element type of a typed-array view — discriminates the 11
/// ECMA-262 §23.2 typed-array variants. Determines both the bytes
/// per element and the sign-extension / clamping / float-conversion
/// applied when reading or writing elements.
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum TypedElemKind {
    /// Int8Array — signed 8-bit; get sign-extends
    I8,
    /// Uint8Array — unsigned 8-bit; get zero-extends
    U8,
    /// Uint8ClampedArray — unsigned 8-bit; set clamps to [0, 255]
    U8Clamped,
    /// Int16Array
    I16,
    /// Uint16Array
    U16,
    /// Int32Array
    I32,
    /// Uint32Array — stored as i32; unsigned interpretation at language boundary
    U32,
    /// Float32Array
    F32,
    /// Float64Array
    F64,
    /// BigInt64Array — elements are i64
    BigI64,
    /// BigUint64Array — stored as i64; unsigned interpretation at language boundary
    BigU64,
}

impl TypedElemKind {
    pub fn bytes_per_element(self) -> usize {
        match self {
            TypedElemKind::I8 | TypedElemKind::U8 | TypedElemKind::U8Clamped => 1,
            TypedElemKind::I16 | TypedElemKind::U16 => 2,
            TypedElemKind::I32 | TypedElemKind::U32 | TypedElemKind::F32 => 4,
            TypedElemKind::F64 | TypedElemKind::BigI64 | TypedElemKind::BigU64 => 8,
        }
    }
}

/// A `TypedArray` view over an `ArrayBuffer`. `buffer` is the shared
/// byte storage from the underlying `ObjectKind::ArrayBuffer`;
/// writes through this view are observable from any other view
/// (`DataView` or other `TypedArray`) on the same buffer.
///
/// `length` is in **elements**, not bytes. `byte_offset` is the view's
/// starting byte within the buffer.
#[derive(Debug, Clone)]
pub struct TypedArrayState {
    pub elem: TypedElemKind,
    pub buffer: Arc<Mutex<Vec<u8>>>,
    /// Reference to the owning ArrayBuffer object, kept so property
    /// accessors (`.buffer`) return the same externref the user
    /// created the view from.
    pub buffer_obj: Arc<Mutex<Object>>,
    pub byte_offset: usize,
    pub length: usize,
}

#[derive(Debug, Clone)]
pub enum ObjectKind {
    /// Plain JS `Object` — property-bag (via the enclosing
    /// `Object::properties` HashMap). Also the fallback for any value
    /// shape that doesn't warrant a dedicated variant.
    Ordinary,
    /// Dense integer-indexed array — JS `Array`, Python `list`, Ruby
    /// `Array`, Dart `List`, VB `ReDim` array, C# `List<T>`, COBOL
    /// `OCCURS DEPENDING ON`.
    Array(Vec<Value>),
    /// Insertion-ordered key→value map with O(1) lookup — JS `Map`,
    /// Python `dict`, Ruby `Hash`, Dart `Map`, C# `Dictionary`,
    /// PHP `array` (PHP uses this *via* `ObjectKind::Ordinary` for
    /// JS-object semantics; when strictly a JS `Map` the compiler
    /// picks this variant instead).
    ///
    /// Dynamic-runtime Phase B4: replaces the previous
    /// tagged-property-bag MVP. Keys use `SameValueZero` semantics
    /// via `Value`'s `Hash + Eq` impls.
    Map(indexmap::IndexMap<Value, Value>),
    /// Insertion-ordered set of unique values with O(1) membership —
    /// JS `Set`, Python `set`, Ruby `Set`, Dart `Set`, C# `HashSet`.
    /// Members use SameValueZero equality (see `Map` above).
    Set(indexmap::IndexSet<Value>),
    /// Raw byte buffer — JS `ArrayBuffer`, Python `bytes` / `bytearray`
    /// backing. Packed `Vec<u8>` gives us 8× memory density over the
    /// previous "Vec of `Value::I32` boxed bytes" MVP and lets
    /// TypedArray views re-interpret bytes at native speed.
    ArrayBuffer(ArrayBufferState),
    /// View over an `ArrayBuffer` — JS `Int8Array` / `Uint8Array` /
    /// `Uint8ClampedArray` / `Int16Array` / `Uint16Array` /
    /// `Int32Array` / `Uint32Array` / `Float32Array` /
    /// `Float64Array` / `BigInt64Array` / `BigUint64Array` per
    /// ECMA-262 §23.2.
    ///
    /// Writes through this view mutate the shared buffer bytes;
    /// other views of the same buffer see the change immediately.
    TypedArray(TypedArrayState),
    Function(Function),
    /// A reference to a host function by its index in the VM's host_fns table.
    HostFunction(usize),
    /// ECMA-262 §10.4.6 Module Namespace Exotic Object — the runtime
    /// materialization of `import * as ns from "wasi:foo"` when read as
    /// a bare value (reflective access: `Object.keys(ns)`,
    /// `Reflect.ownKeys(ns)`, `typeof ns`, `ns[Symbol.toStringTag]`).
    ///
    /// Spec invariants enforced by `host_imports::install` at VM setup
    /// and respected by the VM's property-access path:
    ///   - `[[Prototype]] = null` (no prototype chain walk)
    ///   - `[[Extensible]] = false` (frozen — no new exports at runtime)
    ///   - Own keys = sorted exports ∪ `@@toStringTag`
    ///   - Each export: `[[Enumerable]] = true`, `[[Writable]] = false`,
    ///     `[[Configurable]] = false`
    ///   - `@@toStringTag` = `"Module"`, non-writable / non-enumerable /
    ///     non-configurable
    ///
    /// Hot-path qualified access (`ns.field(args)`) is resolved at
    /// compile time by the Linker — this object is only materialized
    /// when code actually reads `ns` as a value.
    ModuleNamespace,
    /// WASM stack-switching continuation — a coroutine. Holds the entry
    /// function (to call on first resume) and the saved fiber state
    /// (when paused). Each suspend captures the current VM state into
    /// `saved`; each resume either calls `entry` (fresh) or restores
    /// `saved` (paused).
    Continuation(ContinuationState),
    /// CM3 / WASI 0.3 future<T> — single-value async result.
    /// `id` indexes into the EventLoop's future registry.
    /// Awaited via FUTURE_AWAIT opcode; resolved/rejected by host via HostContext.
    Future {
        id: u64,
    },
    /// CM3 / WASI 0.3 stream<T> — async sequence of values.
    /// `id` indexes into the EventLoop's stream registry.
    /// Read via STREAM_READ opcode; pushed/closed by host via HostContext.
    Stream {
        id: u64,
    },
}

/// Runtime state for an `ObjectKind::Continuation`. Tracks the entry
/// function, the fiber captured mid-suspend, and the lifecycle state.
#[derive(Debug)]
pub struct ContinuationState {
    /// Function the continuation wraps. Called once on the first
    /// resume; after that, the coroutine lives entirely in `saved`.
    pub entry: Value,
    /// Mid-suspend fiber — stack + frames + open upvalues. `None` when
    /// the continuation has never run or has finished.
    pub saved: std::sync::Mutex<Option<crate::fiber::Fiber>>,
    /// Lifecycle. `ready` = never resumed; `suspended` = paused
    /// mid-execution; `done` = entry returned normally, no further
    /// resumes allowed.
    pub state: std::sync::Mutex<ContinuationPhase>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContinuationPhase {
    Ready,
    Suspended,
    Done,
}

impl Clone for ContinuationState {
    /// Continuations are identity-like — cloning produces an
    /// independent skeleton with the same entry but a fresh saved
    /// slot. In practice `ObjectKind` cloning happens rarely for
    /// continuations and any caller that does it gets a logically
    /// fresh coroutine.
    fn clone(&self) -> Self {
        ContinuationState {
            entry: self.entry.clone(),
            saved: std::sync::Mutex::new(None),
            state: std::sync::Mutex::new(ContinuationPhase::Ready),
        }
    }
}

impl Object {
    pub fn new() -> Self {
        Object {
            properties: HashMap::new(),
            kind: ObjectKind::Ordinary,
            type_id: 0,
            fields: Vec::new(),
        }
    }

    pub fn new_typed(type_id: usize) -> Self {
        Object {
            properties: HashMap::new(),
            kind: ObjectKind::Ordinary,
            type_id,
            fields: Vec::new(),
        }
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
        obj.properties
            .insert("length".into(), Value::F64(len as f64));
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
                self.properties
                    .insert("length".into(), Value::F64(elems.len() as f64));
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
