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

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::{EventLoop, Task};
use crate::fiber::{Fiber, SavedFrame};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;
use crate::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation, Value};
use crate::vm::{
    VM, CallFrame, ExceptionHandler, FinalizerEntry, LabelEntry,
    ExecResult, HostContext, HostFn, ImportTarget,
    MAX_FRAMES, MAX_STACK,
};

impl VM {
    pub(crate) fn simd_i32x4_binop(&mut self, f: impl Fn(i32, i32) -> i32) -> Result<(), VMError> {
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
    pub(crate) fn simd_f64x2_binop(&mut self, f: impl Fn(f64, f64) -> f64) -> Result<(), VMError> {
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
    pub(crate) fn simd_f64x2_cmp(&mut self, f: impl Fn(f64, f64) -> bool) -> Result<(), VMError> {
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
    pub(crate) fn simd_f32x4_binop(&mut self, f: impl Fn(f32, f32) -> f32) -> Result<(), VMError> {
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
    pub(crate) fn simd_i8x16_binop(&mut self, f: impl Fn(u8, u8) -> u8) -> Result<(), VMError> {
        let b = self.pop(); let a = self.pop();
        if let (Value::V128(va), Value::V128(vb)) = (a, b) {
            let mut out = [0u8; 16];
            for i in 0..16 { out[i] = f(va[i], vb[i]); }
            self.push(Value::V128(out))
        } else { self.push(Value::V128([0; 16])) }
    }
    pub(crate) fn simd_i16x8_binop(&mut self, f: impl Fn(i16, i16) -> i16) -> Result<(), VMError> {
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
    pub(crate) fn test_type(&self, val: &Value, target_name: &str) -> bool {
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
            // Symbols and BigInts never participate in GC-type / inheritance
            // type tests — they're JS primitives.
            Value::Symbol(_) | Value::BigInt(_) => {
                target_name.eq_ignore_ascii_case(val.type_tag())
            }
        }
    }

    /// Check if an exception value matches a tag name.
    /// Works for: string exceptions (by content), objects with __type or __exception_type,
    /// and cross-language name matching (e.g., "ValueError", "TypeError").
    pub(crate) fn exception_value_matches(&self, val: &Value, tag_name: &str) -> bool {
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

}
