//! SIMD binop helpers + runtime type/exception-tag matching used by
//! the opcode dispatcher.
//!
//! * `simd_*_binop` — factored helpers for i32x4 / f64x2 / f32x4 / i8x16
//!   / i16x8 element-wise binary operations. Every SIMD opcode handler
//!   that reduces to "zip two lane vectors through a Rust closure" uses
//!   one of these to avoid repeating stack-pop / splat / push.
//! * `test_type` / `exception_value_matches` — reflective helpers used
//!   by `ref.test`-shaped opcodes and the exception-handling proposal's
//!   typed-catch arms.

use crate::error::VMError;
use crate::value::Value;
use crate::vm::VM;

impl VM {
    pub(crate) fn simd_i32x4_binop(&mut self, f: impl Fn(i32, i32) -> i32) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = i32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                let lb = i32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                out[i * 4..i * 4 + 4].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f64x2_binop(&mut self, f: impl Fn(f64, f64) -> f64) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                let lb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                out[i * 8..i * 8 + 8].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f64x2_cmp(&mut self, f: impl Fn(f64, f64) -> bool) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                let lb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                let mask: u64 = if f(la, lb) { u64::MAX } else { 0 };
                out[i * 8..i * 8 + 8].copy_from_slice(&mask.to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f32x4_binop(&mut self, f: impl Fn(f32, f32) -> f32) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                let lb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                out[i * 4..i * 4 + 4].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i8x16_binop(&mut self, f: impl Fn(u8, u8) -> u8) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..16 {
                out[i] = f(va[i], vb[i]);
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i16x8_binop(&mut self, f: impl Fn(i16, i16) -> i16) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..8 {
                let la = i16::from_le_bytes([va[i * 2], va[i * 2 + 1]]);
                let lb = i16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]);
                out[i * 2..i * 2 + 2].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i8x16_unop(&mut self, f: impl Fn(u8) -> u8) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..16 {
                out[i] = f(a[i]);
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i16x8_unop(&mut self, f: impl Fn(i16) -> i16) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..8 {
                let v = i16::from_le_bytes([a[i * 2], a[i * 2 + 1]]);
                out[i * 2..i * 2 + 2].copy_from_slice(&f(v).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i32x4_unop(&mut self, f: impl Fn(i32) -> i32) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let v = i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap());
                out[i * 4..i * 4 + 4].copy_from_slice(&f(v).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i64x2_unop(&mut self, f: impl Fn(i64) -> i64) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let v = i64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap());
                out[i * 8..i * 8 + 8].copy_from_slice(&f(v).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f32x4_unop(&mut self, f: impl Fn(f32) -> f32) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let v = f32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap());
                out[i * 4..i * 4 + 4].copy_from_slice(&f(v).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f64x2_unop(&mut self, f: impl Fn(f64) -> f64) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let v = f64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap());
                out[i * 8..i * 8 + 8].copy_from_slice(&f(v).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i64x2_binop(&mut self, f: impl Fn(i64, i64) -> i64) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = i64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                let lb = i64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                out[i * 8..i * 8 + 8].copy_from_slice(&f(la, lb).to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_f32x4_cmp(&mut self, f: impl Fn(f32, f32) -> bool) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                let lb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                let mask: u32 = if f(la, lb) { u32::MAX } else { 0 };
                out[i * 4..i * 4 + 4].copy_from_slice(&mask.to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i64x2_cmp(&mut self, f: impl Fn(i64, i64) -> bool) -> Result<(), VMError> {
        let b = self.pop();
        let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = i64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                let lb = i64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                let mask: u64 = if f(la, lb) { u64::MAX } else { 0 };
                out[i * 8..i * 8 + 8].copy_from_slice(&mask.to_le_bytes());
            }
            self.push(Value::V128(out))
        } else {
            self.push(Value::V128([0; 16]))
        }
    }
    pub(crate) fn simd_i8x16_testop(&mut self, f: impl Fn(u8) -> bool) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let result = a.iter().all(|&b| f(b));
            self.push(Value::I32(if result { 1 } else { 0 }))
        } else {
            self.push(Value::I32(0))
        }
    }
    pub(crate) fn simd_i16x8_testop(&mut self, f: impl Fn(i16) -> bool) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let result = (0..8).all(|i| f(i16::from_le_bytes([a[i * 2], a[i * 2 + 1]])));
            self.push(Value::I32(if result { 1 } else { 0 }))
        } else {
            self.push(Value::I32(0))
        }
    }
    pub(crate) fn simd_i32x4_testop(&mut self, f: impl Fn(i32) -> bool) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let result =
                (0..4).all(|i| f(i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap())));
            self.push(Value::I32(if result { 1 } else { 0 }))
        } else {
            self.push(Value::I32(0))
        }
    }
    pub(crate) fn simd_i64x2_testop(&mut self, f: impl Fn(i64) -> bool) -> Result<(), VMError> {
        if let Value::V128(a) = self.pop() {
            let result =
                (0..2).all(|i| f(i64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap())));
            self.push(Value::I32(if result { 1 } else { 0 }))
        } else {
            self.push(Value::I32(0))
        }
    }

    /// Test if a value matches a type name (used by ref_test, ref_cast, br_on_cast).
    /// Supports: WASM GC type_id lookup, __type string matching, __types array
    /// (JS class inheritance chain), and __control_type for GUI controls.
    /// §7.3.19 OrdinaryHasInstance, name-free: is `<target>.prototype` in
    /// `val`'s `__proto__` chain, by object identity? Resolves the target's
    /// prototype ONLY from the `__ctor_<target>` anchor — the convention the
    /// JS prelude wires for its canonical constructors. Deliberately does NOT
    /// fall back to a bare `<target>` global: a language whose classes are
    /// bare globals (e.g. PHP) carries its real hierarchy in the type registry
    /// + `__types`, and a bare-global lookup there mis-resolves sibling error
    /// classes (PHP `Error`/`Exception` both implement `Throwable`). No anchor
    /// ⇒ no-op, and the registry/stamp checks in `test_type` stand.
    pub(crate) fn proto_chain_has(&self, val: &Value, target_name: &str) -> bool {
        let target_proto = self
            .globals
            .get(&format!("__ctor_{target_name}"))
            .and_then(|ctor| {
                if let Value::Object(c) = ctor {
                    c.lock().unwrap().properties.get("prototype").cloned()
                } else {
                    None
                }
            });
        let Some(Value::Object(target_proto)) = target_proto else {
            return false;
        };
        let mut current = match val {
            Value::Object(o) => o.lock().unwrap().properties.get("__proto__").cloned(),
            _ => return false,
        };
        // Bounded walk — a corrupt cyclic chain must not spin forever.
        for _ in 0..1024 {
            let Some(Value::Object(proto)) = current else {
                return false;
            };
            if std::sync::Arc::ptr_eq(&proto, &target_proto) {
                return true;
            }
            current = proto.lock().unwrap().properties.get("__proto__").cloned();
        }
        false
    }

    pub(crate) fn test_type(&self, val: &Value, target_name: &str) -> bool {
        // ── WASM GC abstract heap types (spec §6.2) ───────────────
        // Bottom types: always false for any non-null value.
        if matches!(target_name, "none" | "nofunc" | "noextern") {
            return matches!(val, Value::Null);
        }
        // `any`: top of internal hierarchy — true for every non-null value.
        if target_name == "any" {
            return !matches!(val, Value::Null | Value::Undefined);
        }
        // `extern`: external references (all JS values are externref in Vybe).
        if target_name == "extern" {
            return !matches!(val, Value::Null);
        }
        // `eq`: types on which ref.eq is allowed — i31 + struct + array.
        if target_name == "eq" {
            return match val {
                Value::I32(_) => true,
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    !matches!(
                        ob.kind,
                        crate::value::ObjectKind::Function(_)
                            | crate::value::ObjectKind::HostFunction(_)
                    )
                }
                _ => false,
            };
        }
        // `i31`: unboxed 31-bit integers.
        if target_name == "i31" {
            return matches!(val, Value::I32(_));
        }
        // `func`: top of function hierarchy — funcref.
        if target_name == "func" || target_name == "function" {
            return match val {
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    matches!(
                        ob.kind,
                        crate::value::ObjectKind::Function(_)
                            | crate::value::ObjectKind::HostFunction(_)
                    )
                }
                _ => false,
            };
        }
        // `struct`: top of struct hierarchy — all non-array, non-func objects.
        if target_name == "struct" {
            return match val {
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    !matches!(
                        ob.kind,
                        crate::value::ObjectKind::Array(_)
                            | crate::value::ObjectKind::Function(_)
                            | crate::value::ObjectKind::HostFunction(_)
                    )
                }
                _ => false,
            };
        }
        // `array`: top of array hierarchy.
        if target_name == "array" {
            return match val {
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    matches!(ob.kind, crate::value::ObjectKind::Array(_))
                }
                _ => false,
            };
        }

        // ── Named / user-defined types ────────────────────────────
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

                let obj_type = ob
                    .properties
                    .get("__type")
                    .map(|v| format!("{}", v))
                    .or_else(|| {
                        ob.properties
                            .get("__control_type")
                            .map(|v| format!("{}", v))
                    })
                    .unwrap_or_default();

                if obj_type == target_name {
                    return true;
                }

                if let Some(tid) = self.type_registry.get_id(&obj_type) {
                    if let Some(target_id) = self.type_registry.get_id(target_name) {
                        if self.type_registry.is_subtype(tid, target_id) {
                            return true;
                        }
                    }
                }

                if let Some(Value::Object(types)) = ob.properties.get("__types") {
                    let t = types.lock().unwrap();
                    if let crate::value::ObjectKind::Array(ref elems) = t.kind {
                        if elems
                            .iter()
                            .any(|e| format!("{}", e).eq_ignore_ascii_case(target_name))
                        {
                            return true;
                        }
                    }
                }

                // §7.3.19 OrdinaryHasInstance: walk the object's `__proto__`
                // chain looking for `<target>.prototype` BY OBJECT IDENTITY —
                // the spec-true, name-free type test. Requires the constructor
                // anchor `__ctor_<target>` (or the bare global `<target>`) to
                // carry a `prototype`; when absent (non-JS profiles, non-class
                // targets) this is a no-op and the stamp checks above stand.
                drop(ob);
                if self.proto_chain_has(val, target_name) {
                    return true;
                }

                target_name.eq_ignore_ascii_case("object")
            }
            Value::String(_) => target_name == "string",
            Value::F64(_) | Value::F32(_) | Value::I32(_) | Value::I64(_) => {
                target_name == "integer" || target_name == "double" || target_name == "number"
            }
            Value::Bool(_) => target_name == "boolean",
            Value::V128(_) => target_name == "v128",
            Value::WeakRef(weak) => {
                if let Some(strong) = weak.upgrade() {
                    self.test_type(&Value::Object(strong), target_name)
                } else {
                    false
                }
            }
            Value::Null | Value::Undefined => false,
            Value::Symbol(_) | Value::BigInt(_) => target_name.eq_ignore_ascii_case(val.type_tag()),
        }
    }

    // -- Execute --
}
