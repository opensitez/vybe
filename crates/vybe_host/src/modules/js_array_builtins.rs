//! # `wasm:js-array` host handlers
//!
//! Native Rust implementations that satisfy the `wasm:js-array.*`
//! imports declared in
//! `crates/vybe_bytecode/src/wasm/js_array_builtins.rs`.
//!
//! On the Vybe VM this file IS the host fast-path — handlers operate
//! directly on the native `Vec<Value>` backing an `ObjectKind::Array`,
//! no indirection through spine-struct WASM bytecode. On v8 / browsers
//! the same interface is satisfied by JS glue (Phase C). On plain
//! wasmtime the polyfill module provides spine-struct implementations.
//!
//! Semantics follow ECMA-262 §23.1 exactly. When in doubt, MDN is the
//! authoritative reference.
//!
//! Marshaling + error-handling contract pinned in
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

/// Shorthand: unwrap `args[idx]` as a JS Array. Returns `None` when
/// the argument isn't an array-kind object. Handlers that require an
/// array trap (per convention class 1) on `None`; handlers that match
/// spec "if not an Array, return something sensible" use the None
/// branch to provide the default.
fn array_of<'a>(args: &'a [Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    match args.get(idx) {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            if matches!(o.kind, ObjectKind::Array(_)) {
                drop(o);
                Some(obj.clone())
            } else {
                None
            }
        }
        _ => None,
    }
}

fn make_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

/// Marker property set by `wasm:js-fixedarray.freeze` to forbid
/// length-changing mutations. Mutators check it and no-op rather
/// than allow the change (spec behavior would be TypeError; we
/// silently no-op until exception dispatch from host handlers is
/// wired — Phase B5 follow-up).
const FROZEN_MARK: &str = "__vybe_frozen";

fn is_frozen(arr: &Arc<Mutex<Object>>) -> bool {
    let o = arr.lock().unwrap();
    o.properties.get(FROZEN_MARK).is_some()
}

/// Keep the array's cached `length` property in sync with the
/// backing vector's length. Every mutator must call this after
/// modifying the vector — JS code reading `.length` does not re-query
/// the Vec; it reads the stored property.
fn sync_length(obj: &mut Object) {
    if let ObjectKind::Array(ref v) = obj.kind {
        let n = v.len();
        obj.properties.insert("length".into(), Value::F64(n as f64));
    }
}

pub fn register(vm: &mut VM) {
    register_constructors(vm);
    register_property_access(vm);
    register_mutators(vm);
    register_non_mutators(vm);
    register_iteration(vm);
}

// ── Constructors ──────────────────────────────────────────────────────

fn register_constructors(vm: &mut VM) {
    // new() -> Array
    vm.register_host_fn(
        "wasm:js-array",
        "new",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| make_array(Vec::new())),
    );

    // newWithLength(n: i32) -> Array (n-element, null-filled)
    vm.register_host_fn(
        "wasm:js-array",
        "newWithLength",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            make_array(vec![Value::Null; n])
        }),
    );

    // of(...values) -> Array
    //
    // Spec: `Array.of(...values)` — variadic. Each positional arg becomes an
    // element. Unlike `new Array(n)` which allocates a length, `Array.of(n)`
    // is always a 1-element array `[n]`.
    vm.register_host_fn(
        "wasm:js-array",
        "of",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            make_array(args.to_vec())
        }),
    );

    // from(src, mapFn?) -> Array
    //
    // Spec: `Array.from(iterable, mapFn)`. Accepts arrays, strings, and
    // array-like objects (anything with `length` and numeric keys). When a
    // `mapFn` is supplied it's invoked as `mapFn(value, index)` via the
    // host's `invoke` callback and the result replaces each element.
    vm.register_host_fn(
        "wasm:js-array",
        "from",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            match args.first() {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = s.kind {
                        out.extend(elems.iter().cloned());
                    } else if let Some(len_val) = s.properties.get("length") {
                        let len = len_val.as_f64().max(0.0) as usize;
                        for i in 0..len {
                            let key = i.to_string();
                            out.push(s.properties.get(&key).cloned().unwrap_or(Value::Undefined));
                        }
                    }
                }
                Some(Value::String(s)) => {
                    for c in s.chars() {
                        out.push(Value::String(Arc::from(c.to_string().as_str())));
                    }
                }
                _ => {}
            }
            if let Some(mapper) = args.get(1) {
                if !matches!(mapper, Value::Null | Value::Undefined) {
                    let mut mapped = Vec::with_capacity(out.len());
                    for (i, v) in out.iter().enumerate() {
                        mapped.push(ctx.invoke(mapper, &[v.clone(), Value::F64(i as f64)]));
                    }
                    out = mapped;
                }
            }
            make_array(out)
        }),
    );

    // fromAsync(asyncIterable, mapFn) -> Promise<Array>
    // Stub for now: returns an empty Array; real impl requires async
    // iteration integration with JSPI. Listed in the import set so
    // compilers emitting calls don't fail to link; behavior will be
    // completed in a later pass.
    vm.register_host_fn(
        "wasm:js-array",
        "fromAsync",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| make_array(Vec::new())),
    );

    // isArray(v) -> i32
    vm.register_host_fn(
        "wasm:js-array",
        "isArray",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if array_of(args, 0).is_some() { 1 } else { 0 })
        }),
    );
}

// ── Property access ────────────────────────────────────────────────────

fn register_property_access(vm: &mut VM) {
    // get(arr_or_obj, key) -> value
    //
    // Primary use is `Array.prototype`-style integer indexing, but this
    // import is also the landing pad for `dict.has` / `hasOwnProperty`
    // / `in` compiled through compiler_common. Accepts plain objects to
    // satisfy those callers without requiring a second import.
    vm.register_host_fn(
        "wasm:js-array",
        "get",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let i = key.as_i32();
                        if i < 0 { return Value::Undefined; }
                        return v.get(i as usize).cloned().unwrap_or(Value::Undefined);
                    }
                    // Polymorphic dispatch on Map — the canonical cross-
                    // language associative type (PHP `['k'=>v]`, Python
                    // dicts, Ruby hashes, JS plain objects). Key is looked
                    // up using Value-level equality (SameValueZero), so
                    // both `$m['foo']` and `$m[$key]` where `$key = 'foo'`
                    // resolve identically.
                    ObjectKind::Map(m) => {
                        let lookup_key = match &key {
                            Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F64(_) => key.clone(),
                            other => Value::String(std::sync::Arc::from(format!("{}", other).as_str())),
                        };
                        if let Some(v) = m.get(&lookup_key) {
                            return v.clone();
                        }
                        // PHP-ish fallback: if caller used a string key like
                        // "0" but the map stores integer keys (or vice
                        // versa), try the coerced form. Only coerces for
                        // purely numeric strings to avoid surprises.
                        if let Value::String(s) = &key {
                            if let Ok(n) = s.parse::<i32>() {
                                if let Some(v) = m.get(&Value::I32(n)) { return v.clone(); }
                            }
                        } else if let Value::I32(n) = &key {
                            if let Some(v) = m.get(&Value::String(std::sync::Arc::from(n.to_string().as_str()))) {
                                return v.clone();
                            }
                        }
                        return Value::Undefined;
                    }
                    _ => {}
                }
                // Plain Object fallback: property lookup. Used by Ordinary
                // objects and by compiler_common's `has` / `in` emitter.
                let key_str = match &key {
                    Value::String(s) => s.to_string(),
                    other => format!("{}", other),
                };
                if let Some(v) = o.properties.get(&key_str) {
                    return v.clone();
                }
                return Value::Undefined;
            }
            Value::Undefined
        }),
    );

    // set(arr_or_obj, key, v) -> () — extends arrays with null-fill when
    // key >= length; stores into plain objects by string key; updates Maps
    // using the canonical Value-keyed IndexMap.
    vm.register_host_fn(
        "wasm:js-array",
        "set",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).cloned().unwrap_or(Value::Undefined);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                match &mut o.kind {
                    ObjectKind::Array(v) => {
                        let i = key.as_i32();
                        if i < 0 {
                            return Value::Null;
                        }
                        let idx = i as usize;
                        while v.len() <= idx {
                            v.push(Value::Null);
                        }
                        v[idx] = val;
                        sync_length(&mut o);
                    }
                    ObjectKind::Map(m) => {
                        let map_key = match &key {
                            Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F64(_) => key.clone(),
                            other => Value::String(std::sync::Arc::from(format!("{}", other).as_str())),
                        };
                        m.insert(map_key, val);
                    }
                    _ => {
                        let key_str = match &key {
                            Value::String(s) => s.to_string(),
                            other => format!("{}", other),
                        };
                        o.properties.insert(key_str, val);
                    }
                }
            }
            Value::Null
        }),
    );

    // length(arr) -> i32 — ECMA-262 §23.1.3.12. Strict Array/TypedArray
    // only; strings use `wasm:js-string.length` per the js-string-builtins
    // proposal. Polymorphic callers (e.g. our `__len__` canonical) must
    // type-dispatch before selecting the import.
    vm.register_host_fn(
        "wasm:js-array",
        "length",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(o)) = args.first() {
                let lock = o.lock().unwrap();
                return match &lock.kind {
                    ObjectKind::Array(v) => Value::I32(v.len() as i32),
                    ObjectKind::Map(m) => Value::I32(m.len() as i32),
                    ObjectKind::TypedArray(t) => Value::I32(t.length as i32),
                    _ => lock.properties.get("length")
                        .map(|v| Value::I32(v.as_i32()))
                        .unwrap_or(Value::I32(0)),
                };
            }
            Value::I32(0)
        }),
    );

    // setLength(arr, n) -> () — truncate or null-fill extend
    vm.register_host_fn(
        "wasm:js-array",
        "setLength",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    v.resize(n, Value::Null);
                }
                sync_length(&mut o);
            }
            Value::Null
        }),
    );

    // at(arr, i) -> value
    //
    // `Array.prototype.at` — negative indices relative to length, undefined
    // when OOB. String `.at()` routes through `wasm:js-value.invokeMethod`.
    vm.register_host_fn(
        "wasm:js-array",
        "at",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let idx = if i < 0 { len + i } else { i };
                    if idx < 0 || idx >= len {
                        return Value::Undefined;
                    }
                    return v.get(idx as usize).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );
}

// ── Mutators ──────────────────────────────────────────────────────────

fn register_mutators(vm: &mut VM) {
    // push(arr, v) -> i32 new_length
    //
    // Guards against frozen arrays: a frozen array's length cannot
    // change, so push is a no-op and returns the current length. The
    // spec says TypeError — we'll upgrade to a throw when host-side
    // exception dispatch lands.
    vm.register_host_fn(
        "wasm:js-array",
        "push",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let val = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        return Value::I32(v.len() as i32);
                    }
                    return Value::I32(0);
                }
                let mut o = arr.lock().unwrap();
                let len = if let ObjectKind::Array(ref mut v) = o.kind {
                    v.push(val);
                    v.len() as i32
                } else {
                    0
                };
                sync_length(&mut o);
                return Value::I32(len);
            }
            Value::I32(0)
        }),
    );

    // pop(arr) -> popped_value (undefined if empty or frozen)
    vm.register_host_fn(
        "wasm:js-array",
        "pop",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) { return Value::Undefined; }
                let mut o = arr.lock().unwrap();
                let popped = if let ObjectKind::Array(ref mut v) = o.kind {
                    v.pop().unwrap_or(Value::Undefined)
                } else {
                    Value::Undefined
                };
                sync_length(&mut o);
                return popped;
            }
            Value::Undefined
        }),
    );

    // shift(arr) -> first_value (undefined if empty or frozen)
    vm.register_host_fn(
        "wasm:js-array",
        "shift",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) { return Value::Undefined; }
                let mut o = arr.lock().unwrap();
                let shifted = if let ObjectKind::Array(ref mut v) = o.kind {
                    if v.is_empty() { Value::Undefined } else { v.remove(0) }
                } else {
                    Value::Undefined
                };
                sync_length(&mut o);
                return shifted;
            }
            Value::Undefined
        }),
    );

    // unshift(arr, v1, v2, ...) -> i32 new_length
    //
    // Spec: inserts all v_i at the head in order, so `[3,4,5].unshift(1,2)`
    // yields `[1,2,3,4,5]`. Frozen arrays just return the current length.
    vm.register_host_fn(
        "wasm:js-array",
        "unshift",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        return Value::I32(v.len() as i32);
                    }
                    return Value::I32(0);
                }
                let mut o = arr.lock().unwrap();
                let len = if let ObjectKind::Array(ref mut v) = o.kind {
                    for (i, val) in args.iter().skip(1).enumerate() {
                        v.insert(i, val.clone());
                    }
                    v.len() as i32
                } else {
                    0
                };
                sync_length(&mut o);
                return Value::I32(len);
            }
            Value::I32(0)
        }),
    );

    // splice(arr, start, deleteCount, ...items) -> deleted_array
    //
    // Spec: items come through as variadic individual args, not a wrapped
    // array. `arr.splice(1, 0, 2, 3)` inserts 2 and 3 at index 1 and
    // deletes 0; args[3..] hold the items.
    vm.register_host_fn(
        "wasm:js-array",
        "splice",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let del = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let items: Vec<Value> = args.iter().skip(3).cloned().collect();
            let mut deleted = Vec::new();
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len();
                    let idx = if start < 0 {
                        ((len as i32) + start).max(0) as usize
                    } else {
                        (start as usize).min(len)
                    };
                    let end = (idx + del).min(len);
                    for _ in idx..end {
                        deleted.push(v.remove(idx));
                    }
                    for (i, val) in items.into_iter().enumerate() {
                        v.insert(idx + i, val);
                    }
                }
                sync_length(&mut o);
            }
            make_array(deleted)
        }),
    );

    // reverse(arr) -> self (in-place)
    vm.register_host_fn(
        "wasm:js-array",
        "reverse",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    v.reverse();
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // sort(arr, compareFn) -> self (in-place) — MVP: compare by stringified value
    // Real callback dispatch requires VM `invoke_callback`; implement
    // in the Phase B5 iterator-helpers pass when we tackle callbacks
    // uniformly.
    vm.register_host_fn(
        "wasm:js-array",
        "sort",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    v.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // fill(arr, value, start, end) -> self
    vm.register_host_fn(
        "wasm:js-array",
        "fill",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let val = args.get(1).cloned().unwrap_or(Value::Null);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    for i in s..e {
                        v[i] = val.clone();
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // copyWithin(arr, target, start, end) -> self
    vm.register_host_fn(
        "wasm:js-array",
        "copyWithin",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let target = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len() as i32;
                    let t = target.max(0).min(len) as usize;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let slice: Vec<Value> = v[s..e].iter().cloned().collect();
                    let max_copy = (len as usize - t).min(slice.len());
                    v[t..t + max_copy].clone_from_slice(&slice[..max_copy]);
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );
}

// ── Non-mutators ──────────────────────────────────────────────────────

fn register_non_mutators(vm: &mut VM) {
    // slice(arr, start, end) -> new_arr
    //
    // Array-only per ECMA-262. String slicing goes through
    // `wasm:js-value.invokeMethod` (or `wasm:js-string.slice` under the
    // js-string-builtins proposal when v8 hosts it).
    vm.register_host_fn(
        "wasm:js-array",
        "slice",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let s = (if start < 0 { len + start } else { start }).max(0).min(len) as usize;
                    let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    let out: Vec<Value> = if s < e { v[s..e].to_vec() } else { Vec::new() };
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );

    // concat(arr, other) -> new_arr
    vm.register_host_fn(
        "wasm:js-array",
        "concat",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    out.extend(v.iter().cloned());
                }
            }
            // Spec: if `other` is an array, spread it; otherwise append as single element
            match args.get(1) {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    match &lock.kind {
                        ObjectKind::Array(v) => out.extend(v.iter().cloned()),
                        _ => out.push(Value::Object(o.clone())),
                    }
                }
                Some(v) => out.push(v.clone()),
                None => {}
            }
            make_array(out)
        }),
    );

    // indexOf(arr, value, fromIndex) -> i32
    vm.register_host_fn(
        "wasm:js-array",
        "indexOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            let from = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let start = from.max(0) as usize;
                    for (i, elem) in v.iter().enumerate().skip(start) {
                        if elem.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    // lastIndexOf(arr, value, fromIndex) -> i32
    vm.register_host_fn(
        "wasm:js-array",
        "lastIndexOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let from = args.get(2).map(|v| v.as_i32()).unwrap_or(len - 1);
                    let end = from.min(len - 1).max(-1);
                    let end_idx = if end < 0 { 0 } else { (end + 1) as usize };
                    for (i, elem) in v[..end_idx].iter().enumerate().rev() {
                        if elem.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    // includes(arr_or_obj, value, fromIndex) -> bool
    //
    // Primary: `Array.prototype.includes` (SameValueZero comparison). Also
    // the landing pad for the compiled `x in y` operator when `y` is a
    // plain object — we check own-property membership for that case.
    // String `.includes(...)` routes through `wasm:js-value.invokeMethod`.
    vm.register_host_fn(
        "wasm:js-array",
        "includes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            let from = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    for elem in v.iter().skip(from) {
                        if elem.eq(&needle) {
                            return Value::Bool(true);
                        }
                    }
                    return Value::Bool(false);
                }
                let key_str = match &needle {
                    Value::String(s) => s.to_string(),
                    other => format!("{}", other),
                };
                return Value::Bool(o.properties.contains_key(&key_str));
            }
            Value::Bool(false)
        }),
    );

    // join(arr, sep) -> string
    vm.register_host_fn(
        "wasm:js-array",
        "join",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let sep = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| ",".into());
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let parts: Vec<String> = v
                        .iter()
                        .map(|e| match e {
                            Value::Null | Value::Undefined => String::new(),
                            _ => format!("{}", e),
                        })
                        .collect();
                    return Value::String(Arc::from(parts.join(&sep).as_str()));
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // toString(arr) -> string (same as join with default ",")
    vm.register_host_fn(
        "wasm:js-array",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let parts: Vec<String> = v
                        .iter()
                        .map(|e| match e {
                            Value::Null | Value::Undefined => String::new(),
                            _ => format!("{}", e),
                        })
                        .collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // toLocaleString — same as toString for MVP
    vm.register_host_fn(
        "wasm:js-array",
        "toLocaleString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            // Same as toString — real locale-aware conversion lives in
            // Phase F (intl integration).
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let parts: Vec<String> = v.iter().map(|e| format!("{}", e)).collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // flat(arr, depth) -> new_arr
    vm.register_host_fn(
        "wasm:js-array",
        "flat",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let depth = args.get(1).map(|v| v.as_i32()).unwrap_or(1);
            fn flatten(out: &mut Vec<Value>, input: &[Value], depth: i32) {
                for v in input {
                    if depth > 0 {
                        if let Value::Object(o) = v {
                            let lock = o.lock().unwrap();
                            if let ObjectKind::Array(ref inner) = lock.kind {
                                flatten(out, inner, depth - 1);
                                continue;
                            }
                        }
                    }
                    out.push(v.clone());
                }
            }
            let mut out = Vec::new();
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    flatten(&mut out, v, depth);
                }
            }
            make_array(out)
        }),
    );

    // ── ES2023 non-mutating variants ────────────────────────────────

    // toReversed(arr) -> new_arr
    vm.register_host_fn(
        "wasm:js-array",
        "toReversed",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut out = v.clone();
                    out.reverse();
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );

    // toSorted(arr, compareFn) -> new_arr
    vm.register_host_fn(
        "wasm:js-array",
        "toSorted",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut out = v.clone();
                    out.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );

    // with(arr, i, v) -> new_arr
    vm.register_host_fn(
        "wasm:js-array",
        "with",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let idx = if i < 0 { len + i } else { i };
                    if idx < 0 || idx >= len {
                        // Per spec: throw RangeError. For MVP we return
                        // the array unchanged — Phase B6 test gate will
                        // catch this and we'll add proper throw.
                        return make_array(v.clone());
                    }
                    let mut out = v.clone();
                    out[idx as usize] = val;
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );
}

// ── Iteration / higher-order callbacks ─────────────────────────────────

fn register_iteration(vm: &mut VM) {
    // keys(arr) / values(arr) / entries(arr) — return Array of keys,
    // values, or [k, v] pairs. Spec returns iterators; Phase B12 will
    // upgrade these to real iterator externrefs. MVP returns an Array
    // which satisfies most callers since arrays are iterable.
    vm.register_host_fn(
        "wasm:js-array",
        "keys",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let out: Vec<Value> = (0..v.len()).map(|i| Value::F64(i as f64)).collect();
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "wasm:js-array",
        "values",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return make_array(v.clone());
                }
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "wasm:js-array",
        "entries",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let out: Vec<Value> = v
                        .iter()
                        .enumerate()
                        .map(|(i, e)| make_array(vec![Value::F64(i as f64), e.clone()]))
                        .collect();
                    return make_array(out);
                }
            }
            make_array(Vec::new())
        }),
    );

    // ── Callback-taking methods (Phase B5 — real dispatch) ────────────
    //
    // These invoke the supplied JS callback per element via
    // `HostContext::invoke`, matching the callback signature
    // `(element, index, array) → result` that the MDN reference
    // pages specify for Array.prototype methods.
    //
    // Spec references (each method):
    //   - forEach: §23.1.3.13 — no return; just invoke for side effects
    //   - map:     §23.1.3.21 — collect invoke results
    //   - filter:  §23.1.3.8  — keep elements where callback is truthy
    //   - reduce:  §23.1.3.26 — fold from left, optional initial value
    //   - reduceRight: §23.1.3.27
    //   - some:    §23.1.3.30 — any callback returns truthy
    //   - every:   §23.1.3.6  — all callbacks return truthy
    //   - find, findLast: §23.1.3.11 — first/last element where truthy
    //   - findIndex, findLastIndex: §23.1.3.12 — index of first/last match
    //   - flatMap: §23.1.3.15 — map + flatten one level

    vm.register_host_fn("wasm:js-array", "forEach",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    ctx.invoke(&callback, &invoke_args);
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("wasm:js-array", "map",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                let mapped: Vec<Value> = snapshot.iter().enumerate()
                    .map(|(i, elem)| {
                        let invoke_args = vec![
                            elem.clone(),
                            Value::I32(i as i32),
                            Value::Object(arr.clone()),
                        ];
                        ctx.invoke(&callback, &invoke_args)
                    })
                    .collect();
                return make_array(mapped);
            }
            make_array(Vec::new())
        }));

    vm.register_host_fn("wasm:js-array", "filter",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                let filtered: Vec<Value> = snapshot.iter().enumerate()
                    .filter_map(|(i, elem)| {
                        let invoke_args = vec![
                            elem.clone(),
                            Value::I32(i as i32),
                            Value::Object(arr.clone()),
                        ];
                        let keep = is_truthy(&ctx.invoke(&callback, &invoke_args));
                        if keep { Some(elem.clone()) } else { None }
                    })
                    .collect();
                return make_array(filtered);
            }
            make_array(Vec::new())
        }));

    vm.register_host_fn("wasm:js-array", "reduce",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let initial_provided = args.len() > 2 && !matches!(args.get(2), Some(Value::Undefined) | None);
            let mut acc = if initial_provided {
                args.get(2).cloned().unwrap_or(Value::Undefined)
            } else {
                Value::Undefined
            };
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                let start_idx = if initial_provided { 0 } else {
                    if snapshot.is_empty() {
                        // Spec: TypeError on empty array with no initial.
                        // MVP returns undefined; Phase B5 doesn't have
                        // throw-dispatch yet.
                        return Value::Undefined;
                    }
                    acc = snapshot[0].clone();
                    1
                };
                for i in start_idx..snapshot.len() {
                    let invoke_args = vec![
                        acc,
                        snapshot[i].clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    acc = ctx.invoke(&callback, &invoke_args);
                }
            }
            acc
        }));

    vm.register_host_fn("wasm:js-array", "reduceRight",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let initial_provided = args.len() > 2 && !matches!(args.get(2), Some(Value::Undefined) | None);
            let mut acc = if initial_provided {
                args.get(2).cloned().unwrap_or(Value::Undefined)
            } else {
                Value::Undefined
            };
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                if snapshot.is_empty() {
                    return if initial_provided { acc } else { Value::Undefined };
                }
                let mut i = snapshot.len() as i32 - 1;
                if !initial_provided {
                    acc = snapshot[i as usize].clone();
                    i -= 1;
                }
                while i >= 0 {
                    let invoke_args = vec![
                        acc,
                        snapshot[i as usize].clone(),
                        Value::I32(i),
                        Value::Object(arr.clone()),
                    ];
                    acc = ctx.invoke(&callback, &invoke_args);
                    i -= 1;
                }
            }
            acc
        }));

    vm.register_host_fn("wasm:js-array", "some",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return Value::I32(1);
                    }
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-array", "every",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if !is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return Value::I32(0);
                    }
                }
            }
            Value::I32(1) // spec: empty array → every returns true
        }));

    vm.register_host_fn("wasm:js-array", "find",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return elem.clone();
                    }
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("wasm:js-array", "findIndex",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn("wasm:js-array", "findLast",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate().rev() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return elem.clone();
                    }
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("wasm:js-array", "findLastIndex",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate().rev() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&ctx.invoke(&callback, &invoke_args)) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn("wasm:js-array", "flatMap",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                let mut out = Vec::with_capacity(snapshot.len());
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    let r = ctx.invoke(&callback, &invoke_args);
                    // Flatten one level: if the result is an Array, spread;
                    // otherwise append as single element.
                    if let Value::Object(ref o) = r {
                        let lock = o.lock().unwrap();
                        if let ObjectKind::Array(ref inner) = lock.kind {
                            out.extend(inner.iter().cloned());
                            continue;
                        }
                    }
                    out.push(r);
                }
                return make_array(out);
            }
            make_array(Vec::new())
        }));

    // ── ES2025 group / groupToMap ───────────────────────────────────
    //
    // Group elements by the result of the callback. `group` returns a
    // null-prototype Object with string keys; `groupToMap` returns a
    // Map keyed by any value.

    vm.register_host_fn("wasm:js-array", "group",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let mut groups: IndexMap<String, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    let key = format!("{}", ctx.invoke(&callback, &invoke_args));
                    groups.entry(key).or_insert_with(Vec::new).push(elem.clone());
                }
            }
            // Materialize as an ordinary object with array-valued properties.
            let mut out = Object::new();
            for (k, v) in groups {
                out.properties.insert(k, make_array(v));
            }
            Value::Object(Arc::new(Mutex::new(out)))
        }));

    vm.register_host_fn("wasm:js-array", "groupToMap",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let mut groups: IndexMap<Value, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    let key = ctx.invoke(&callback, &invoke_args);
                    groups.entry(key).or_insert_with(Vec::new).push(elem.clone());
                }
            }
            // Build a JS Map with one entry per group.
            let mut map_im: IndexMap<Value, Value> = IndexMap::new();
            for (k, v) in groups {
                map_im.insert(k, make_array(v));
            }
            let mut obj = Object::new();
            obj.kind = ObjectKind::Map(map_im);
            obj.properties.insert("size".into(), Value::I32(obj_map_len(&obj) as i32));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // toSpliced — non-mutating splice returning a new array.
    vm.register_host_fn("wasm:js-array", "toSpliced",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let del = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let items: Vec<Value> = match args.get(3) {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref v) = lock.kind { v.clone() } else { Vec::new() }
                }
                _ => Vec::new(),
            };
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind { v.clone() } else { Vec::new() }
                };
                let len = snapshot.len();
                let idx = if start < 0 {
                    ((len as i32) + start).max(0) as usize
                } else {
                    (start as usize).min(len)
                };
                let end = (idx + del).min(len);
                let mut out = Vec::with_capacity(len - (end - idx) + items.len());
                out.extend_from_slice(&snapshot[..idx]);
                out.extend(items.into_iter());
                out.extend_from_slice(&snapshot[end..]);
                return make_array(out);
            }
            make_array(Vec::new())
        }));
}

/// JS truthy semantics — used by filter / some / every / find.
/// Matches ECMA-262 §7.1.2 ToBoolean.
fn is_truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::String(s) => !s.is_empty(),
        Value::Object(_) | Value::Symbol(_) | Value::BigInt(_)
            | Value::V128(_) | Value::WeakRef(_) => true,
    }
}

fn obj_map_len(obj: &Object) -> usize {
    if let ObjectKind::Map(ref m) = obj.kind { m.len() } else { 0 }
}
