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

#[path = "simd_kernels.rs"]
mod kernels;

#[inline(always)]
fn lane_i32(bytes: &[u8; 16], lane: usize) -> i32 {
    let i = lane * 4;
    i32::from_le_bytes([bytes[i], bytes[i + 1], bytes[i + 2], bytes[i + 3]])
}

#[inline(always)]
fn lane_i64(bytes: &[u8; 16], lane: usize) -> i64 {
    let i = lane * 8;
    i64::from_le_bytes([
        bytes[i],
        bytes[i + 1],
        bytes[i + 2],
        bytes[i + 3],
        bytes[i + 4],
        bytes[i + 5],
        bytes[i + 6],
        bytes[i + 7],
    ])
}

#[inline(always)]
fn lane_f32(bytes: &[u8; 16], lane: usize) -> f32 {
    let i = lane * 4;
    f32::from_le_bytes([bytes[i], bytes[i + 1], bytes[i + 2], bytes[i + 3]])
}

#[inline(always)]
fn lane_f64(bytes: &[u8; 16], lane: usize) -> f64 {
    let i = lane * 8;
    f64::from_le_bytes([
        bytes[i],
        bytes[i + 1],
        bytes[i + 2],
        bytes[i + 3],
        bytes[i + 4],
        bytes[i + 5],
        bytes[i + 6],
        bytes[i + 7],
    ])
}

#[inline(always)]
fn write_i16_lane(out: &mut [u8; 16], lane: usize, value: i16) {
    let i = lane * 2;
    let bytes = value.to_le_bytes();
    out[i] = bytes[0];
    out[i + 1] = bytes[1];
}

#[inline(always)]
fn write_i32_lane(out: &mut [u8; 16], lane: usize, value: i32) {
    let i = lane * 4;
    let bytes = value.to_le_bytes();
    out[i] = bytes[0];
    out[i + 1] = bytes[1];
    out[i + 2] = bytes[2];
    out[i + 3] = bytes[3];
}

#[inline(always)]
fn write_u32_lane(out: &mut [u8; 16], lane: usize, value: u32) {
    let i = lane * 4;
    let bytes = value.to_le_bytes();
    out[i] = bytes[0];
    out[i + 1] = bytes[1];
    out[i + 2] = bytes[2];
    out[i + 3] = bytes[3];
}

#[inline(always)]
fn write_i64_lane(out: &mut [u8; 16], lane: usize, value: i64) {
    let i = lane * 8;
    let bytes = value.to_le_bytes();
    out[i] = bytes[0];
    out[i + 1] = bytes[1];
    out[i + 2] = bytes[2];
    out[i + 3] = bytes[3];
    out[i + 4] = bytes[4];
    out[i + 5] = bytes[5];
    out[i + 6] = bytes[6];
    out[i + 7] = bytes[7];
}

#[inline(always)]
fn write_u64_lane(out: &mut [u8; 16], lane: usize, value: u64) {
    let i = lane * 8;
    let bytes = value.to_le_bytes();
    out[i] = bytes[0];
    out[i + 1] = bytes[1];
    out[i + 2] = bytes[2];
    out[i + 3] = bytes[3];
    out[i + 4] = bytes[4];
    out[i + 5] = bytes[5];
    out[i + 6] = bytes[6];
    out[i + 7] = bytes[7];
}

#[inline(always)]
fn write_f32_lane(out: &mut [u8; 16], lane: usize, value: f32) {
    write_u32_lane(out, lane, value.to_bits());
}

#[inline(always)]
fn write_f64_lane(out: &mut [u8; 16], lane: usize, value: f64) {
    write_u64_lane(out, lane, value.to_bits());
}

impl VM {
    /// Keep the result in the lower operand slot. This avoids two pops,
    /// a push-capacity check, and intermediate vector copies.
    #[inline(always)]
    fn simd_replace_pair(
        &mut self,
        f: impl FnOnce(&[u8; 16], &[u8; 16]) -> [u8; 16],
    ) -> Result<(), VMError> {
        let len = self.stack.len();
        assert!(len >= 2, "stack underflow");
        let out = match (&self.stack[len - 2], &self.stack[len - 1]) {
            (Value::V128(a), Value::V128(b)) => f(a, b),
            _ => [0; 16],
        };
        self.stack[len - 2] = Value::V128(out);
        self.stack.truncate(len - 1);
        Ok(())
    }

    #[inline(always)]
    fn simd_replace_unary(
        &mut self,
        f: impl FnOnce(&[u8; 16]) -> Value,
        fallback: Value,
    ) -> Result<(), VMError> {
        let out = match self.stack.last().expect("stack underflow") {
            Value::V128(a) => f(a),
            _ => fallback,
        };
        *self.stack.last_mut().expect("stack underflow") = out;
        Ok(())
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_eq(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_eq)
    }

    #[inline(always)]
    pub(crate) fn simd_i32x4_eq(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i32x4_eq)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_ne(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_ne)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_eq(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_eq)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_add(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_add)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_swizzle(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_swizzle)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_add_sat_s(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_add_sat_s)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_add_sat_u(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_add_sat_u)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_sub_sat_s(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_sub_sat_s)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_sub_sat_u(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_sub_sat_u)
    }

    #[inline(always)]
    pub(crate) fn simd_i8x16_sub(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i8x16_sub)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_add(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_add)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_avgr_u(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_avgr_u)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_bitmask(&mut self) -> Result<(), VMError> {
        self.simd_replace_unary(|bytes| Value::I32(kernels::i16x8_bitmask(bytes)), Value::I32(0))
    }

    #[cfg(test)]
    #[inline(always)]
    pub(crate) fn simd_i16x8_all_true(&mut self) -> Result<(), VMError> {
        self.simd_replace_unary(|bytes| Value::I32(kernels::i16x8_all_true(bytes)), Value::I32(0))
    }

    #[cfg(test)]
    #[inline(always)]
    fn simd_i16x8_all_true_scalar_reference(&mut self) -> Result<(), VMError> {
        self.simd_i16x8_testop(|lane| lane != 0)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_add_sat_s(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_add_sat_s)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_add_sat_u(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_add_sat_u)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_sub_sat_s(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_sub_sat_s)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_sub_sat_u(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_sub_sat_u)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_sub(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_sub)
    }

    #[inline(always)]
    pub(crate) fn simd_i16x8_mul(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i16x8_mul)
    }

    #[inline(always)]
    pub(crate) fn simd_i32x4_add(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i32x4_add)
    }

    #[inline(always)]
    pub(crate) fn simd_i32x4_sub(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i32x4_sub)
    }

    #[inline(always)]
    pub(crate) fn simd_i64x2_add(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i64x2_add)
    }

    #[inline(always)]
    pub(crate) fn simd_i64x2_sub(&mut self) -> Result<(), VMError> {
        self.simd_replace_pair(kernels::i64x2_sub)
    }
    #[inline(always)]
    pub(crate) fn simd_i32x4_binop(&mut self, f: impl Fn(i32, i32) -> i32) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = lane_i32(va, i);
                let lb = lane_i32(vb, i);
                write_i32_lane(&mut out, i, f(la, lb));
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_f64x2_binop(&mut self, f: impl Fn(f64, f64) -> f64) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = lane_f64(va, i);
                let lb = lane_f64(vb, i);
                write_f64_lane(&mut out, i, f(la, lb));
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_f64x2_cmp(&mut self, f: impl Fn(f64, f64) -> bool) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = lane_f64(va, i);
                let lb = lane_f64(vb, i);
                let mask: u64 = if f(la, lb) { u64::MAX } else { 0 };
                write_u64_lane(&mut out, i, mask);
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_f32x4_binop(&mut self, f: impl Fn(f32, f32) -> f32) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = lane_f32(va, i);
                let lb = lane_f32(vb, i);
                write_f32_lane(&mut out, i, f(la, lb));
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_i8x16_binop(&mut self, f: impl Fn(u8, u8) -> u8) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..16 {
                out[i] = f(va[i], vb[i]);
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_i16x8_binop(&mut self, f: impl Fn(i16, i16) -> i16) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..8 {
                let la = i16::from_le_bytes([va[i * 2], va[i * 2 + 1]]);
                let lb = i16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]);
                write_i16_lane(&mut out, i, f(la, lb));
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_i8x16_unop(&mut self, f: impl Fn(u8) -> u8) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..16 {
                out[i] = f(a[i]);
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_i16x8_unop(&mut self, f: impl Fn(i16) -> i16) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..8 {
                let v = i16::from_le_bytes([a[i * 2], a[i * 2 + 1]]);
                write_i16_lane(&mut out, i, f(v));
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_i32x4_unop(&mut self, f: impl Fn(i32) -> i32) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let v = lane_i32(a, i);
                write_i32_lane(&mut out, i, f(v));
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_i64x2_unop(&mut self, f: impl Fn(i64) -> i64) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let v = lane_i64(a, i);
                write_i64_lane(&mut out, i, f(v));
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_f32x4_unop(&mut self, f: impl Fn(f32) -> f32) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let v = lane_f32(a, i);
                write_f32_lane(&mut out, i, f(v));
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_f64x2_unop(&mut self, f: impl Fn(f64) -> f64) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let v = lane_f64(a, i);
                write_f64_lane(&mut out, i, f(v));
            }
            Value::V128(out)
        }, Value::V128([0; 16]))
    }
    #[inline(always)]
    pub(crate) fn simd_i64x2_binop(&mut self, f: impl Fn(i64, i64) -> i64) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = lane_i64(va, i);
                let lb = lane_i64(vb, i);
                write_i64_lane(&mut out, i, f(la, lb));
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_f32x4_cmp(&mut self, f: impl Fn(f32, f32) -> bool) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..4 {
                let la = lane_f32(va, i);
                let lb = lane_f32(vb, i);
                let mask: u32 = if f(la, lb) { u32::MAX } else { 0 };
                write_u32_lane(&mut out, i, mask);
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_i64x2_cmp(&mut self, f: impl Fn(i64, i64) -> bool) -> Result<(), VMError> {
        self.simd_replace_pair(|va, vb| {
            let mut out = [0u8; 16];
            for i in 0..2 {
                let la = lane_i64(va, i);
                let lb = lane_i64(vb, i);
                let mask: u64 = if f(la, lb) { u64::MAX } else { 0 };
                write_u64_lane(&mut out, i, mask);
            }
            out
        })
    }
    #[inline(always)]
    pub(crate) fn simd_i8x16_testop(&mut self, f: impl Fn(u8) -> bool) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let result = a.iter().all(|&b| f(b));
            Value::I32(if result { 1 } else { 0 })
        }, Value::I32(0))
    }
    #[inline(always)]
    pub(crate) fn simd_i16x8_testop(&mut self, f: impl Fn(i16) -> bool) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let result = (0..8).all(|i| f(i16::from_le_bytes([a[i * 2], a[i * 2 + 1]])));
            Value::I32(if result { 1 } else { 0 })
        }, Value::I32(0))
    }
    #[inline(always)]
    pub(crate) fn simd_i32x4_testop(&mut self, f: impl Fn(i32) -> bool) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let result = (0..4).all(|i| f(lane_i32(a, i)));
            Value::I32(if result { 1 } else { 0 })
        }, Value::I32(0))
    }
    #[inline(always)]
    pub(crate) fn simd_i64x2_testop(&mut self, f: impl Fn(i64) -> bool) -> Result<(), VMError> {
        self.simd_replace_unary(|a| {
            let result = (0..2).all(|i| f(lane_i64(a, i)));
            Value::I32(if result { 1 } else { 0 })
        }, Value::I32(0))
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
            .global(&format!("__ctor_{target_name}"))
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

    /// `ref.test` — the WASM operation. Takes a heap type, never a name.
    pub(crate) fn ref_test(&self, val: &Value, ht: crate::opcode::heaptype::HeapType) -> bool {
        match ht {
            crate::opcode::heaptype::HeapType::Abstract(a) => self.test_abstract(val, a),
            crate::opcode::heaptype::HeapType::Concrete(idx) => self.test_concrete(val, idx),
        }
    }

    /// `ref.test` against an ABSTRACT heap type (GC proposal §6.2).
    ///
    /// A pure shape test on the value: the abstract hierarchy is fixed by the
    /// spec, so nothing here consults the type registry, and there is nothing
    /// to look up by name.
    pub(crate) fn test_abstract(&self, val: &Value, ht: u8) -> bool {
        use crate::opcode::heaptype::*;
        use crate::value::ObjectKind;
        // Which branch of the internal hierarchy an object sits in. The three
        // are disjoint, which is what makes `struct` "neither array nor func".
        let is_array = |val: &Value| matches!(val, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Array(_)));
        let is_func = |val: &Value| {
            matches!(val, Value::Object(o) if matches!(
                o.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
        };
        match ht {
            // Bottom types: inhabited only by null.
            HT_NONE | HT_NOFUNC | HT_NOEXTERN => val.is_null_ref(),
            // `any`: top of the internal hierarchy.
            HT_ANY => !val.is_null_ref() && !matches!(val, Value::Undefined),
            // `extern`: external references (all JS values are externref here).
            HT_EXTERN => !val.is_null_ref(),
            // `eq`: the types `ref.eq` is allowed on — i31 + struct + array.
            HT_EQ => {
                matches!(val, Value::I32(_)) || (matches!(val, Value::Object(_)) && !is_func(val))
            }
            // `i31`: unboxed 31-bit integers.
            HT_I31 => matches!(val, Value::I32(_)),
            // `func`: top of the function hierarchy.
            HT_FUNC => is_func(val),
            // `struct`: every non-array, non-func object.
            HT_STRUCT => matches!(val, Value::Object(_)) && !is_array(val) && !is_func(val),
            // `array`: top of the array hierarchy.
            HT_ARRAY => is_array(val),
            _ => false,
        }
    }

    /// The declared signature of a module type index, when that type is a
    /// `(func …)`. `None` for struct/array types, which are matched by rtt.
    /// The signature lives in the entry's `fields` — the same general payload
    /// an array type's element storage type uses.
    fn declared_func_sig(&self, type_index: u32) -> Option<(&str, &str)> {
        let idx = type_index as usize;
        if idx == 0 {
            return None;
        }
        let entry = self.chunks.first()?.types.get(idx - 1)?;
        if entry.kind != crate::chunk::CompositeKind::Func {
            return None;
        }
        Some((
            entry.fields.first().map(String::as_str).unwrap_or_default(),
            entry.fields.get(1).map(String::as_str).unwrap_or_default(),
        ))
    }

    /// The chunk a callable value refers to, for values that are functions.
    fn function_chunk_index(&self, val: &Value) -> Option<usize> {
        match val {
            Value::Object(o) => match &o.lock().unwrap().kind {
                crate::value::ObjectKind::Function(f) => Some(f.chunk_index),
                _ => None,
            },
            _ => None,
        }
    }

    /// `ref.test` against a CONCRETE type — a module type index, resolved the
    /// same way `struct.new`'s immediate is. An index walk over declared
    /// supertypes; no name, no fallback. A value with no rtt is simply not an
    /// instance of a declared type.
    pub(crate) fn test_concrete(&self, val: &Value, type_index: u32) -> bool {
        // Named heap types are struct/array/function references. Scalar and
        // null values cannot match one, so reject them before module/type
        // lookup. Object references still take the complete subtype/signature
        // path below; nullable tests handle null before reaching this method.
        if !matches!(val, Value::Object(_)) {
            return false;
        }
        let target = self.resolve_gc_rtt(type_index as usize);
        if let Some(ty) = self.type_registry.get(target).filter(|ty| ty.native_identity.is_some()) {
            if ty.kind == crate::chunk::CompositeKind::Func {
                return self.function_chunk_index(val)
                    .and_then(|index| self.chunks.get(index)?.native_func_type.as_ref())
                    .and_then(|identity| self.type_registry.get_native_id(identity))
                    .is_some_and(|actual| self.type_registry.is_subtype(actual, target));
            }
            return match val {
                Value::Object(object) => {
                    let actual = object.lock().unwrap().type_id;
                    actual != 0 && self.type_registry.is_subtype(actual, target)
                }
                _ => false,
            };
        }
        // A FUNCTION reference is matched STRUCTURALLY, not by rtt: a function
        // carries no allocation-time type stamp, and the spec's rule
        // (`Comptype_sub/func`) compares the parameter and result TYPES with
        // no name anywhere in it. Two identically-shaped types declared under
        // different names therefore match, and `(func (param i32))` does not
        // match `(func (param f64))` even though both are 1→0.
        if let Some((want_params, want_results)) = self.declared_func_sig(type_index) {
            return match self.function_chunk_index(val) {
                Some(ci) => match &self.chunks[ci].func_sig {
                    Some((params, results)) => params.as_str() == want_params && results.as_str() == want_results,
                    // A function with no recorded signature cannot be claimed
                    // to match — answering true here would be a guess.
                    None => false,
                },
                None => false,
            };
        }
        if target == 0 {
            return false;
        }
        match val {
            Value::Object(o) => {
                let type_id = o.lock().unwrap().type_id;
                type_id > 0 && self.type_registry.is_subtype(type_id, target)
            }
            _ => false,
        }
    }

    /// The name a module declared for a concrete heap type, for diagnostics
    /// and for the transitional fallback below. Abstract types are spelled by
    /// the spec.
    pub(crate) fn heaptype_label(&self, ht: crate::opcode::heaptype::HeapType) -> String {
        use crate::opcode::heaptype::*;
        match ht {
            HeapType::Abstract(byte) => match byte {
                HT_ANY => "any",
                HT_EQ => "eq",
                HT_I31 => "i31",
                HT_STRUCT => "struct",
                HT_ARRAY => "array",
                HT_FUNC => "func",
                HT_EXTERN => "extern",
                HT_NONE => "none",
                HT_NOFUNC => "nofunc",
                HT_NOEXTERN => "noextern",
                _ => "?",
            }
            .to_string(),
            HeapType::Concrete(index) => self.declared_type_name(index).unwrap_or_default(),
        }
    }

    /// The name the module's type section gives a concrete type index.
    pub(crate) fn declared_type_name(&self, type_index: u32) -> Option<String> {
        if type_index == 0 {
            return None;
        }
        let base = self
            .frames
            .last()
            .and_then(|frame| self.chunk_type_base.get(frame.chunk_index))
            .copied()
            .unwrap_or(0);
        self.module_type_names
            .get(base + type_index as usize - 1)
            .cloned()
    }

    /// `ref.test`, plus the transitional fallback for objects that never got
    /// an rtt.
    ///
    /// The 184 platform stamps across 59 files still identify their objects
    /// with a `__type` STRING rather than a typed allocation, so a concrete
    /// test that finds no rtt has to consult them or those objects become
    /// untypeable. The name for that comes from the module's **type section** —
    /// a declaration — never from the instruction, which carries only an index.
    /// Deleting this is its own step and needs the platforms to gain a typed
    /// allocation path first.
    /// `ref.test`/`ref.cast` against an EXACT heap type — `(ref (exact $t))`.
    ///
    /// The difference from the inexact form is one word: the reference's own
    /// type must BE `$t`, not merely be a subtype of it. `exact-casts.wast`
    /// pins both directions — casting a `$super` to `(exact $super)` returns
    /// it, casting a `$sub` to `(exact $super)` traps — and the subtype walk
    /// answers `true` to both.
    ///
    /// Abstract heap types are not handled specially because the binary format
    /// cannot spell one: `heaptype ::= 0x62 x:u32` takes a plain u32, which the
    /// proposal notes "intentionally makes it impossible to encode an exact
    /// abstract heap type". An abstract immediate here can only come from a
    /// malformed input, so it falls back to the ordinary test rather than
    /// inventing an answer.
    /// Are two DECLARED function types the same type, for an exact cast?
    ///
    /// ⚠ NOT a name comparison. WASM type identity is STRUCTURAL after
    /// canonicalisation, and our type names are module-qualified — so module A's
    /// `(func)` and module B's `(func)` have different NAMES and are the same
    /// TYPE. `exact-func-import.wast` asserts exactly that: a function imported
    /// from another module still casts to the importer's `(ref (exact $f))`.
    ///
    /// But the structure that matters includes the SUBTYPE DECLARATION, not
    /// just params and results. That is what separates `exact-casts.wast`'s
    ///   (type $super (sub (func (result funcref))))
    ///   (type $sub   (sub $super (func (result funcref))))
    /// which share a signature and are still distinct types — `$sub` declares a
    /// supertype and `$super` does not. Comparing signatures alone would call
    /// them identical and let a `$sub` pass an exact cast to `$super`;
    /// comparing NAMES alone would call A's `(func)` and B's `(func)` distinct
    /// and fail a legal cross-module cast. Both halves are needed.
    fn func_types_identical(&self, a: &str, b: &str) -> bool {
        if a == b {
            return true;
        }
        let Some(types) = self.chunks.first().map(|c| &c.types) else {
            return false;
        };
        let entry = |n: &str| types.iter().find(|t| t.name == n);
        let (Some(ea), Some(eb)) = (entry(a), entry(b)) else {
            return false;
        };
        if ea.fields != eb.fields {
            return false;
        }
        // Supertype declarations must correspond. Comparing the parents'
        // SHAPES rather than their indices keeps this module-independent for
        // the same reason the signature comparison is.
        let parent_shape = |e: &crate::chunk::TypeEntry| -> Option<Vec<String>> {
            if e.parent_index == 0 {
                None
            } else {
                types
                    .get(e.parent_index as usize - 1)
                    .map(|p| p.fields.clone())
            }
        };
        parent_shape(ea) == parent_shape(eb)
    }

    pub(crate) fn ref_test_exact(
        &self,
        val: &Value,
        ht: crate::opcode::heaptype::HeapType,
    ) -> bool {
        let crate::opcode::heaptype::HeapType::Concrete(index) = ht else {
            return self.ref_test(val, ht);
        };
        let target = self.resolve_gc_rtt(index as usize);
        if let Some(ty) = self.type_registry.get(target).filter(|ty| ty.native_identity.is_some()) {
            if ty.kind == crate::chunk::CompositeKind::Func {
                return self.function_chunk_index(val)
                    .and_then(|index| self.chunks.get(index)?.native_func_type.as_ref())
                    .and_then(|identity| self.type_registry.get_native_id(identity))
                    .is_some_and(|actual| actual == target);
            }
            return matches!(val, Value::Object(object) if object.lock().unwrap().type_id == target);
        }
        // ⚠ A FUNCTION TYPE IS MATCHED NOMINALLY WHEN THE CAST IS EXACT.
        //
        // Ordinary `ref.test`/`ref.cast` compare function types STRUCTURALLY —
        // `Comptype_sub/func` names no type, only params and results. Exactness
        // cannot be expressed that way: `exact-casts.wast` declares
        //   (type $super (sub (func (result funcref))))
        //   (type $sub   (sub $super (func (result funcref))))
        // — identical structures, distinct names — and asserts that casting a
        // `$sub` function to `(ref (exact $super))` TRAPS. Structure says they
        // match; only the declared NAME tells them apart.
        //
        // Falls back to the structural answer when the function carries no
        // declared type (nothing to compare), rather than inventing a verdict.
        if self.declared_func_sig(index).is_some() {
            let want = self.declared_type_name(index);
            let got = self
                .function_chunk_index(val)
                .and_then(|ci| self.chunks[ci].declared_func_type.clone());
            return match (want, got) {
                (Some(w), Some(g)) => self.func_types_identical(&w, &g),
                _ => self.ref_test(val, ht),
            };
        }
        // ⚠ RESOLVE THE RTT THE WAY `test_concrete` DOES — `resolve_gc_rtt`,
        // which is module-relative — NOT by looking the declared NAME up in the
        // registry. The name route disagrees for array types (`(array i8)`
        // reached the registry under a different key), so every exact cast in
        // `exact-casts.wast`'s array module failed a cast that should succeed.
        // One resolution path for both tests is the point.
        if target == 0 {
            return false;
        }
        match val {
            // The ONLY difference from `test_concrete`: `==` rather than
            // `is_subtype`. That is what "exact" means.
            Value::Object(o) => {
                let type_id = o.lock().unwrap().type_id;
                type_id > 0 && type_id == target
            }
            _ => false,
        }
    }



    /// Type test **by name** — a LANGUAGE operation, not a WASM one.
    ///
    /// PHP's `$x instanceof $className` with a runtime string, JS `instanceof`
    /// against a non-class callable, and every platform object identified only
    /// by a `__type` stamp land here. Kept deliberately separate from
    /// [`ref_test`] so that a gap in the type registry cannot be silently
    /// papered over by a string comparison: if a caller reaches this, it is
    /// because it genuinely had a name and nothing else.
    pub(crate) fn test_type(&self, val: &Value, target_name: &str) -> bool {
        // The spec's own spellings still answer from the abstract hierarchy —
        // `wast` and the reader both name them this way.
        if let Some(ht) = crate::opcode::heaptype::HeapType::from_spec_name(target_name) {
            return self.ref_test(val, ht);
        }
        // `function` is JS's spelling of `func`.
        if target_name == "function" {
            return self.test_abstract(val, crate::opcode::heaptype::HT_FUNC);
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
            Value::Null | Value::TypedNull(_) | Value::Undefined => false,
            Value::Symbol(_) | Value::BigInt(_) => target_name.eq_ignore_ascii_case(val.type_tag()),
        }
    }

    // -- Execute --
}

#[cfg(test)]
#[path = "simd_perf_tests.rs"]
mod bulk_simd_perf_gates;

impl crate::vm::VM {
    #[inline(always)]
    pub(crate) fn simd_i8x16_bitmask_native(&mut self) -> Result<(), crate::error::VMError> {
        let operand = self.stack.last_mut().expect("stack underflow");
        let mask = match operand {
            crate::value::Value::V128(bytes) => kernels::i8x16_bitmask(bytes),
            _ => 0,
        };
        *operand = crate::value::Value::I32(mask);
        Ok(())
    }
}

#[cfg(test)]
mod all_true_stack_tests {
    use super::*;
    use std::{hint::black_box, time::Instant};

    #[test]
    fn i16_all_true_replaces_only_top_operand_with_canonical_i32() {
        for bytes in [[0u8; 16], [1u8; 16], [0x80u8; 16]] {
            let mut vm = VM::new();
            vm.stack = vec![Value::I64(-7), Value::V128(bytes)];
            vm.simd_i16x8_all_true().unwrap();
            assert_eq!(vm.stack.len(), 2);
            assert_eq!(vm.stack[0], Value::I64(-7));
            let expected = bytes.chunks_exact(2)
                .all(|lane| lane[0] != 0 || lane[1] != 0) as i32;
            assert_eq!(vm.stack[1], Value::I32(expected));
        }
    }

    fn scalar(vm: &mut VM) {
        vm.simd_i16x8_all_true_scalar_reference().unwrap();
    }

    fn optimized(vm: &mut VM) {
        vm.simd_i16x8_all_true().unwrap();
    }

    fn sample(kernel: fn(&mut VM), vm: &mut VM, inputs: &[[u8; 16]]) -> u128 {
        let kernel = black_box(kernel);
        let start = Instant::now();
        for index in 0..100_000 {
            vm.stack[1] = Value::V128(inputs[index % inputs.len()]);
            kernel(black_box(&mut *vm));
            black_box(&vm.stack[1]);
        }
        start.elapsed().as_nanos()
    }

    #[test]
    #[ignore = "paired performance experiment; run optimized test executable explicitly"]
    fn benchmark_i16_all_true_actual_stack_helpers() {
        assert!(!cfg!(debug_assertions), "compile with optimizations enabled");
        let mut seed = 0x41b8a531u32;
        let mut mixed = Vec::with_capacity(512);
        for index in 0..512 {
            let mut bytes = [0u8; 16];
            for byte in &mut bytes {
                seed ^= seed << 13;
                seed ^= seed >> 17;
                seed ^= seed << 5;
                *byte = seed as u8;
            }
            if index % 3 != 0 {
                let offset = (index % 8) * 2;
                bytes[offset..offset + 2].fill(0);
            }
            mixed.push(bytes);
        }
        for (pattern, inputs) in [
            ("mixed", mixed),
            ("nonzero", vec![[1u8; 16]; 512]),
            ("zero", vec![[0u8; 16]; 512]),
        ] {
            let mut vm = VM::new();
            vm.stack = vec![Value::I64(-7), Value::V128([0; 16])];
            let mut ratios = Vec::with_capacity(21);
            for round in 0..23 {
                let (before, after) = if round % 2 == 0 {
                    (sample(scalar, &mut vm, &inputs), sample(optimized, &mut vm, &inputs))
                } else {
                    let after = sample(optimized, &mut vm, &inputs);
                    (sample(scalar, &mut vm, &inputs), after)
                };
                assert_eq!(vm.stack.len(), 2);
                assert_eq!(vm.stack[0], Value::I64(-7));
                let bytes = inputs[99_999 % inputs.len()];
                let expected = bytes.chunks_exact(2)
                    .all(|lane| lane[0] != 0 || lane[1] != 0) as i32;
                assert_eq!(vm.stack[1], Value::I32(expected));
                if round >= 2 { ratios.push(before as f64 / after as f64); }
            }
            ratios.sort_by(f64::total_cmp);
            println!("i16x8.all_true stack pattern={pattern} paired_ratio={:.3}", ratios[10]);
        }
    }
}
