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
    fn declared_func_sig(&self, type_index: u32) -> Option<(String, String)> {
        let idx = type_index as usize;
        if idx == 0 {
            return None;
        }
        let entry = self.chunks.first()?.types.get(idx - 1)?;
        if entry.kind != crate::chunk::CompositeKind::Func {
            return None;
        }
        Some((
            entry.fields.first().cloned().unwrap_or_default(),
            entry.fields.get(1).cloned().unwrap_or_default(),
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
        // A FUNCTION reference is matched STRUCTURALLY, not by rtt: a function
        // carries no allocation-time type stamp, and the spec's rule
        // (`Comptype_sub/func`) compares the parameter and result TYPES with
        // no name anywhere in it. Two identically-shaped types declared under
        // different names therefore match, and `(func (param i32))` does not
        // match `(func (param f64))` even though both are 1→0.
        if let Some((want_params, want_results)) = self.declared_func_sig(type_index) {
            return match self.function_chunk_index(val) {
                Some(ci) => match &self.chunks[ci].func_sig {
                    Some((params, results)) => {
                        *params == want_params && *results == want_results
                    }
                    // A function with no recorded signature cannot be claimed
                    // to match — answering true here would be a guess.
                    None => false,
                },
                None => false,
            };
        }
        let target = self.resolve_gc_rtt(type_index as usize);
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
                types.get(e.parent_index as usize - 1).map(|p| p.fields.clone())
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
        let target = self.resolve_gc_rtt(index as usize);
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

    /// Does this value carry a real rtt — i.e. was it allocated as a declared
    /// type rather than as a dynamic object?
    ///
    /// `type_id == 0` is the untyped form (`struct.new 0`), which is what
    /// platform exceptions and object literals still get.
    pub(crate) fn value_has_rtt(&self, val: &Value) -> bool {
        match val {
            Value::Object(o) => o.lock().unwrap().type_id > 0,
            _ => false,
        }
    }

    pub(crate) fn ref_test_or_declared_name(
        &self,
        val: &Value,
        ht: crate::opcode::heaptype::HeapType,
    ) -> bool {
        if self.ref_test(val, ht) {
            return true;
        }
        match ht {
            crate::opcode::heaptype::HeapType::Abstract(_) => false,
            // A concrete FUNCTION type is decided structurally and that answer
            // is FINAL — falling through to the name path would let a
            // `__type`/prototype match override it, so a cast to a function
            // type the reference does not have would succeed instead of
            // trapping. The name path exists for language-level type tests,
            // which have nothing to say about a WASM function signature.
            crate::opcode::heaptype::HeapType::Concrete(index)
                if self.declared_func_sig(index).is_some() =>
            {
                false
            }
            // ⛔⛔ THE NAME PATH IS FOR OBJECTS WITH NO rtt — NEVER AS AN
            // OVERRIDE OF ONE.
            //
            // `test_type` answers from `properties["__type"]` / `["__types"]`,
            // which are ORDINARY WRITABLE PROPERTIES. Consulting them for an
            // object that already carries an rtt made identity FORGEABLE:
            //
            //     const u = new User();
            //     u.__type = "Admin";
            //     u instanceof Admin   →  true      ⛔
            //
            // A real instance re-labelled itself into another class by
            // assignment, and every consumer that believes it is asking the rtt
            // — typed `catch`, `instanceof`, `is`, `isinstance`, the seam-3
            // receiver guard, the private-field brand check — was really asking
            // "does this object carry the right string".
            //
            // The rtt is stamped at allocation and there is no instruction that
            // can change it, so for a typed object `ref_test` is already the
            // complete and unforgeable answer. The name path exists only for
            // values that never got an rtt — platform exceptions and dynamic
            // object literals, both of which are still allocated untyped — and
            // it is scoped to exactly those here rather than trusted globally.
            //
            // ⚠ This does NOT make identity unforgeable for an untyped object;
            // nothing can until every allocation carries an rtt. It removes the
            // case where a GENUINE instance can be relabelled, which is the one
            // that reads as privilege escalation.
            crate::opcode::heaptype::HeapType::Concrete(index)
                if self.value_has_rtt(val) =>
            {
                false
            }
            crate::opcode::heaptype::HeapType::Concrete(index) => self
                .declared_type_name(index)
                .is_some_and(|name| self.test_type(val, &name)),
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
