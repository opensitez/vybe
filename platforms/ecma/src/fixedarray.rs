//! # `ecma:fixedarray` host handlers
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
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

/// Property marker denoting "this array's length is immutable".
/// Mutation handlers (push/pop/shift/splice/setLength) check this
/// and return/trap if set.
const FROZEN_MARK: &str = "__vybe_frozen";

fn fixed_length(value: &Value) -> Value {
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(values) = &o.kind {
            return Value::I32(values.len() as i32);
        }
    }
    Value::I32(0)
}

fn fixed_get(value: &Value, index: i32) -> Value {
    if index >= 0 {
        if let Value::Object(obj) = value {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(values) = &o.kind {
                return values.get(index as usize).cloned().unwrap_or(Value::Undefined);
            }
        }
    }
    Value::Undefined
}

fn fixed_marker(value: &Value) -> Value {
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        if matches!(&o.kind, ObjectKind::Array(_)) {
            return Value::I32(i32::from(o.properties.contains_key(FROZEN_MARK)));
        }
    }
    Value::I32(0)
}

fn freeze_array(value: &Value) -> Value {
    if let Value::Object(obj) = value {
        let mut o = obj.lock().unwrap();
        if matches!(&o.kind, ObjectKind::Array(_)) {
            o.properties.insert(FROZEN_MARK.into(), Value::I32(1));
            drop(o);
            return Value::Object(obj.clone());
        }
    }
    value.clone()
}

pub fn register(vm: &mut VM) {
    // newWithLength(n) — null-filled, frozen.
    vm.register_host_fn(
        "ecma:fixedarray",
        "newWithLength",
        Box::new(|_ctx, args| {
            let n = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            let mut obj = Object::new_array(vec![Value::Null; n]);
            obj.properties.insert(FROZEN_MARK.into(), Value::I32(1));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // fromArray(src) — snapshot elements into a new frozen Array.
    vm.register_host_fn(
        "ecma:fixedarray",
        "fromArray",
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
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // toArray(fixed) — unfreeze copy into a new growable Array.
    vm.register_host_fn(
        "ecma:fixedarray",
        "toArray",
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
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(elements)))
        }),
    );

    vm.register_host_fn(
        "ecma:fixedarray",
        "length",
        Box::new(|_ctx, args| fixed_length(args.first().unwrap_or(&Value::Null))),
    );

    vm.register_host_fn(
        "ecma:fixedarray",
        "get",
        Box::new(|_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            fixed_get(args.first().unwrap_or(&Value::Null), i)
        }),
    );

    vm.register_host_fn(
        "ecma:fixedarray",
        "isFixedArray",
        Box::new(|_ctx, args| fixed_marker(args.first().unwrap_or(&Value::Null))),
    );

    // freeze(arr) — mark an existing growable Array as fixed. Returns
    // the same array (frozen in place).
    vm.register_host_fn(
        "ecma:fixedarray",
        "freeze",
        Box::new(|_ctx, args| freeze_array(args.first().unwrap_or(&Value::Null))),
    );

    vm.register_host_fn(
        "ecma:fixedarray",
        "isFrozen",
        Box::new(|_ctx, args| fixed_marker(args.first().unwrap_or(&Value::Null))),
    );
}

#[cfg(test)]
mod array_lock_speedup_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn array(values: Vec<Value>) -> Value {
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
    }

    #[test]
    fn reads_preserve_length_bounds_identity_and_marker_presence() {
        let element = vybe_runtime::heap::alloc(Object::new());
        let value = array(vec![Value::I32(7), Value::Undefined, Value::Object(element.clone())]);
        assert_eq!(fixed_length(&value).as_i32(), 3);
        assert_eq!(fixed_get(&value, 0).as_i32(), 7);
        assert!(matches!(fixed_get(&value, -1), Value::Undefined));
        assert!(matches!(fixed_get(&value, 1), Value::Undefined));
        assert!(matches!(fixed_get(&value, 3), Value::Undefined));
        assert!(matches!(fixed_get(&value, 2), Value::Object(actual) if Arc::ptr_eq(&actual, &element)));
        assert_eq!(fixed_marker(&value).as_i32(), 0);
        let Value::Object(obj) = &value else { panic!("expected array") };
        obj.lock().unwrap().properties.insert(FROZEN_MARK.into(), Value::I32(0));
        assert_eq!(fixed_marker(&value).as_i32(), 1);
        let frozen = freeze_array(&value);
        assert!(matches!(frozen, Value::Object(actual) if Arc::ptr_eq(&actual, obj)));
        assert!(obj.try_lock().is_ok());
        assert_eq!(fixed_length(&value).as_i32(), 3);
    }

    #[test]
    fn invalid_receivers_and_empty_arrays_keep_existing_results() {
        let mut obj = Object::new();
        obj.properties.insert(FROZEN_MARK.into(), Value::I32(1));
        let ordinary = Value::Object(vybe_runtime::heap::alloc(obj));
        for value in [Value::Null, Value::Undefined, Value::I32(9), ordinary] {
            assert_eq!(fixed_length(&value).as_i32(), 0);
            assert_eq!(fixed_marker(&value).as_i32(), 0);
            assert!(matches!(fixed_get(&value, 0), Value::Undefined));
            let returned = freeze_array(&value);
            match (&value, &returned) {
                (Value::Object(a), Value::Object(b)) => assert!(Arc::ptr_eq(a, b)),
                _ => assert_eq!(value, returned),
            }
        }
        let empty = array(Vec::new());
        assert_eq!(fixed_length(&empty).as_i32(), 0);
        assert!(matches!(fixed_get(&empty, 0), Value::Undefined));
    }

    fn original_length(value: &Value) -> Value {
        let arr = if let Value::Object(obj) = value {
            let o = obj.lock().unwrap();
            if matches!(&o.kind, ObjectKind::Array(_)) {
                drop(o);
                Some(obj.clone())
            } else { None }
        } else { None };
        if let Some(arr) = arr {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(values) = &o.kind {
                return Value::I32(values.len() as i32);
            }
        }
        Value::I32(0)
    }

    #[test]
    #[ignore = "native fixed-array read microbenchmark, not complete VM dispatch timing"]
    fn native_length_read_microbenchmark() {
        let value = array(vec![Value::Null; 128]);
        let mut old = Duration::ZERO;
        let mut new = Duration::ZERO;
        for batch in 0..10 {
            for old_path in if batch % 2 == 0 { [true, false] } else { [false, true] } {
                let start = Instant::now();
                for _ in 0..10000 {
                    let length = if old_path { original_length(black_box(&value)) }
                        else { fixed_length(black_box(&value)) };
                    assert_eq!(black_box(length).as_i32(), 128);
                }
                if old_path { old += start.elapsed(); } else { new += start.elapsed(); }
            }
        }
        eprintln!("native FixedArray length, calls=100000: validate-clone-relock={old:?}, borrowed-single-lock={new:?}");
    }
}
