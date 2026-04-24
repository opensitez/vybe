//! # `vybe:js-fixedarray` host handlers
//!
//! FixedArray is not a separate `ObjectKind` variant — it's a
//! *compile-time intent* layered on top of `ObjectKind::Array`:
//! an array with the `__vybe_frozen` marker property set, which
//! the growable-array handlers check before mutating.
//!
//! Rationale: the underlying storage is the same `Vec<Value>` either
//! way; the only difference is whether `push`/`pop`/`shift`/`splice`
//! are permitted. A runtime flag is cheaper than a dedicated enum
//! variant when the operations are identical up to mutation gating,
//! and it lets compilers seamlessly promote an Array to a
//! FixedArray via `freeze()` without copying.
//!
//! COBOL `OCCURS n TIMES DEPENDING ON v` → `Array`
//! COBOL `OCCURS n TIMES`                 → `FixedArray` (Array + frozen)
//! Python `tuple`                          → `FixedArray`
//! VB `Dim arr(5)` without `ReDim`         → `FixedArray`
//! C# `T[]` when compiler proves static    → `FixedArray`
//!
//! See `JS_BUILTIN_CONVENTIONS.md`.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

/// Property marker denoting "this array's length is immutable".
/// Mutation handlers (push/pop/shift/splice/setLength) check this
/// and return/trap if set.
const FROZEN_MARK: &str = "__vybe_frozen";

fn array_of(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Array(_)) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    // newWithLength(n) — null-filled, frozen.
    vm.register_host_fn("vybe:js-fixedarray", "newWithLength",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let mut obj = Object::new_array(vec![Value::Null; n]);
            obj.properties.insert(FROZEN_MARK.into(), Value::I32(1));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // fromArray(src) — snapshot elements into a new frozen Array.
    vm.register_host_fn("vybe:js-fixedarray", "fromArray",
        Box::new(|_ctx, args| {
            let elements: Vec<Value> = match args.first() {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref v) = s.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            let mut obj = Object::new_array(elements);
            obj.properties.insert(FROZEN_MARK.into(), Value::I32(1));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // toArray(fixed) — unfreeze copy into a new growable Array.
    vm.register_host_fn("vybe:js-fixedarray", "toArray",
        Box::new(|_ctx, args| {
            let elements: Vec<Value> = match args.first() {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref v) = s.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            // No FROZEN_MARK — growable.
            Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
        }));

    vm.register_host_fn("vybe:js-fixedarray", "length",
        Box::new(|_ctx, args| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return Value::I32(v.len() as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-fixedarray", "get",
        Box::new(|_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return Value::Undefined; }
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return v.get(i as usize).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("vybe:js-fixedarray", "isFixedArray",
        Box::new(|_ctx, args| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                return Value::I32(
                    if o.properties.get(FROZEN_MARK).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    // freeze(arr) — mark an existing growable Array as fixed. Returns
    // the same array (frozen in place).
    vm.register_host_fn("vybe:js-fixedarray", "freeze",
        Box::new(|_ctx, args| {
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                o.properties.insert(FROZEN_MARK.into(), Value::I32(1));
                drop(o);
                return Value::Object(arr);
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn("vybe:js-fixedarray", "isFrozen",
        Box::new(|_ctx, args| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                return Value::I32(
                    if o.properties.get(FROZEN_MARK).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));
}
