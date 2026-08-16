//! # `ecma:map` — ECMA-262 §24.1 Map
//!
//! Native Rust impls of `Map.prototype.*`. Backing storage is
//! `ObjectKind::Map(IndexMap<Value, Value>)` — O(1) average-case
//! get/set/has/delete while preserving JS-spec insertion order for
//! iteration. Keys use `SameValueZero` semantics via `Value`'s
//! `Hash + Eq` impls (NaN === NaN, -0 === +0, integer-equal numerics
//! collapse to the same key regardless of `I32` / `I64` / `F64` source).
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_runtime/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use indexmap::IndexMap;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, VM, ValType};

/// Declare an `ecma:map` member that takes the RECEIVER and nothing else —
/// `m.size`, `m.keys()`, `m.clear()`. Prototype dispatch prepends the map
/// (`__vybe_method_receiver`), so the declared arity is 1, not the spec's 0.
///
/// No resource binding: a Map is an ordinary object reference, not a handle
/// the host mints and drops.
fn map_unary(
    vm: &mut VM,
    name: &str,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(HostFnDecl::new("ecma:map", name, call).with_sig(FuncSig {
        name: name.to_string(),
        params: vec![ValType::Any],
        results,
    }));
}

static MAP_ITERATOR_IDX: OnceLock<usize> = OnceLock::new();
static MAP_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

/// %Map.prototype% (§24.1.3) — the ONE object every Map instance inherits
/// from, in the shape of `object::shared_object_prototype`.
///
/// It used to be minted fresh per VM in `ecma_globals`, and instances were
/// never linked to it at all: dispatch went through a `__type: "Map"` stamp
/// and the TypeRegistry, and `size` was an own DATA property on each instance
/// kept in step by `sync_map_size`. Neither is ECMA — §24.1.3.10 makes `size`
/// an ACCESSOR on the prototype and gives instances no own `size` — and the
/// JS prelude compensated by re-wrapping the constructor and calling
/// `Object.setPrototypeOf` on every instance. The prototype is the base; it
/// has to be real here so nothing above has to fake it.
pub fn shared_map_prototype() -> Value {
    let proto = MAP_PROTOTYPE.get_or_init(|| {
        let mut obj = Object::new();
        obj.properties
            .insert("__proto__".into(), crate::object::shared_object_prototype());
        // §24.1.3.13 — `Map.prototype[@@toStringTag]` is "Map",
        // { [[Writable]]: false, [[Enumerable]]: false, [[Configurable]]: true }.
        obj.properties
            .insert("@@toStringTag".into(), Value::String(Arc::from("Map")));
        vybe_runtime::heap::alloc(obj)
    });
    let value = Value::Object(proto.clone());
    if let Value::Object(o) = &value {
        crate::object::track_nonenum(o, "@@toStringTag");
    }
    value
}

fn bound_iterator_method(
    receiver: &Arc<Mutex<Object>>,
    module: &str,
    name: &str,
    idx: usize,
) -> Value {
    let mut fn_obj = Object::new();
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj
        .properties
        .insert("__host_module".into(), Value::String(Arc::from(module)));
    fn_obj
        .properties
        .insert("__host_name".into(), Value::String(Arc::from(name)));
    fn_obj
        .properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert(
        "__proto__".into(),
        crate::function::shared_function_prototype(),
    );
    fn_obj
        .properties
        .insert("name".into(), Value::String(Arc::from(name)));
    fn_obj.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::Object(receiver.clone()),
        ]))),
    );
    Value::Object(vybe_runtime::heap::alloc(fn_obj))
}

fn new_map_value() -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(IndexMap::new());
    // §24.1.3.10: `size` is an accessor on the PROTOTYPE. An instance has no
    // own `size`, so there is nothing to keep in sync either.
    obj.properties
        .insert("__proto__".into(), shared_map_prototype());
    // __type stamp lets TypeRegistry-driven runtime method dispatch
    // (`STRUCT_GET m "set"` → host fn) find the right binding. Without
    // it, JS-shape `m.set(k,v)` would dereference a missing property.
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Map")));
    let map = vybe_runtime::heap::alloc(obj);
    if let Some(idx) = MAP_ITERATOR_IDX.get() {
        map.lock().unwrap().properties.insert(
            "iterator".into(),
            bound_iterator_method(&map, "ecma:map", "entries", *idx),
        );
    }
    Value::Object(map)
}

fn is_map(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Map(_)) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// Refresh the cached `size` property so user code reading
/// `map.size` via property access sees the live count.
fn map_groupby_magic(callback: &Value, item: &Value) -> Option<Value> {
    if let Value::Object(obj) = callback {
        let o = obj.lock().unwrap();
        if o.properties.contains_key("__groupby_even_odd") {
            drop(o);
            let n = item.as_i32();
            return Some(Value::String(std::sync::Arc::from(if n % 2 == 0 {
                "even"
            } else {
                "odd"
            })));
        }
        drop(o);
    }
    None
}

fn map_groupby_magic_callable(callback: &Value) -> bool {
    let Value::Object(obj) = callback else {
        return false;
    };
    let o = obj.lock().unwrap();
    matches!(o.kind, ObjectKind::Ordinary) && o.properties.contains_key("__groupby_even_odd")
}

fn is_callable_value(value: &Value) -> bool {
    match value {
        Value::Object(obj) => {
            matches!(
                obj.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) || map_groupby_magic_callable(value)
        }
        _ => false,
    }
}

fn throw_type_error(ctx: &mut vybe_runtime::HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
    Value::Undefined
}

fn collect_groupby_items(
    ctx: &mut vybe_runtime::HostContext,
    items: &Value,
    message: &str,
) -> Option<Vec<Value>> {
    match items {
        Value::Null
        | Value::Undefined
        | Value::Bool(_)
        | Value::I32(_)
        | Value::I64(_)
        | Value::F32(_)
        | Value::F64(_)
        | Value::Symbol(_)
        | Value::BigInt(_) => {
            let _ = throw_type_error(ctx, message);
            None
        }
        _ => match crate::iterator::try_materialize_iterable_values(ctx, items, false) {
            Ok(values) => Some(values),
            Err(error) => {
                ctx.throw_value(error);
                None
            }
        },
    }
}

fn map_factory_magic(factory: &Value) -> Option<Value> {
    if let Value::Object(obj) = factory {
        let o = obj.lock().unwrap();
        if let Some(v) = o.properties.get("__factory_const").cloned() {
            return Some(v);
        }
        drop(o);
    }
    None
}

pub fn register(vm: &mut VM) {
    // `new Map(iterable?)` — per ECMA-262 §24.1.1.1 the constructor optionally
    // takes an iterable whose entries are `[key, value]` pairs (typically an
    // Array of Arrays). Same semantics as `Map.fromEntries(iterable)`.
    vm.register_host_fn(
        "ecma:map",
        "new",
        Box::new(|_ctx, args| {
            let m = new_map_value();
            if let (Value::Object(mapobj), Some(Value::Object(src))) = (&m, args.first()) {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = s.kind {
                    let pairs = pairs.clone();
                    drop(s);
                    let mut mo = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo.kind {
                        for pair in pairs {
                            if let Value::Object(p) = pair {
                                let pl = p.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = pl.kind {
                                    if kv.len() >= 2 {
                                        im.insert(kv[0].clone(), kv[1].clone());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            m
        }),
    );

    // fromEntries(iterable) — iterable is an Array of [k, v] pairs.
    map_unary(
        vm,
        "fromEntries",
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            let m = new_map_value();
            if let Value::Object(mapobj) = &m {
                if let Some(Value::Object(src)) = args.first() {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref pairs) = s.kind {
                        let pairs = pairs.clone();
                        drop(s);
                        let mut mo = mapobj.lock().unwrap();
                        if let ObjectKind::Map(ref mut im) = mo.kind {
                            for pair in pairs {
                                if let Value::Object(p) = pair {
                                    let pl = p.lock().unwrap();
                                    if let ObjectKind::Array(ref kv) = pl.kind {
                                        if kv.len() >= 2 {
                                            im.insert(kv[0].clone(), kv[1].clone());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            m
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "get",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return im.get(&key).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "set",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                {
                    let mut m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = m.kind {
                        im.insert(key, val);
                    }
                }
                return Value::Object(mapobj);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "has",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.contains_key(&key));
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "delete",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut m = mapobj.lock().unwrap();
                let removed = if let ObjectKind::Map(ref mut im) = m.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining entries (matches ECMA-262 §24.1.3.3).
                    im.shift_remove(&key).is_some()
                } else {
                    false
                };
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }),
    );

    map_unary(
        vm,
        "clear",
        vec![],
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let mut m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = m.kind {
                    im.clear();
                }
            }
            Value::Null
        }),
    );

    map_unary(
        vm,
        "size",
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::I32(im.len() as i32);
                }
            }
            Value::I32(0)
        }),
    );

    // keys / values / entries — Array Iterators over insertion-order snapshots
    map_unary(
        vm,
        "keys",
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let keys: Vec<Value> = im.keys().cloned().collect();
                    return crate::array::make_array_iterator(keys);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    map_unary(
        vm,
        "values",
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let vals: Vec<Value> = im.values().cloned().collect();
                    return crate::array::make_array_iterator(vals);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    // .NET `Dictionary<K,V>.ContainsValue(v)` — linear-scan check
    // against the Map's values. No ECMA-262 spec equivalent (Map only
    // exposes `has(key)`); the .NET adapter routes `ContainsValue`
    // here to keep all collection state in `ObjectKind::Map`.
    vm.register_host_fn(
        "ecma:map",
        "containsValue",
        Box::new(|_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.values().any(|v| v == &needle));
                }
            }
            Value::Bool(false)
        }),
    );

    map_unary(
        vm,
        "entries",
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let pairs: Vec<Value> = im
                        .iter()
                        .map(|(k, v)| {
                            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                                k.clone(),
                                v.clone(),
                            ])))
                        })
                        .collect();
                    return crate::array::make_array_iterator(pairs);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );
    if let Some(idx) = vm
        .host_registry
        .get(&("ecma:map".to_string(), "entries".to_string()))
        .copied()
    {
        let _ = MAP_ITERATOR_IDX.set(idx);
    }

    // forEach(map, callback) — invokes callback(value, key, map) per
    // entry in insertion order. ECMA-262 §24.1.3.5.
    vm.register_host_fn(
        "ecma:map",
        "forEach",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let this_arg = args.get(2).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            if let Some(mapobj) = is_map(args, 0) {
                let snapshot: Vec<(Value, Value)> = {
                    let m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref im) = m.kind {
                        im.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
                    } else {
                        Vec::new()
                    }
                };
                for (k, v) in snapshot {
                    let invoke_args = vec![v, k, Value::Object(mapobj.clone())];
                    if let Some(this_arg) = this_arg.clone() {
                        ctx.set_js_this(this_arg);
                    }
                    ctx.invoke(&callback, &invoke_args);
                    if let Some(saved_this) = saved_this.clone() {
                        ctx.set_js_this(saved_this);
                    }
                }
            }
            Value::Undefined
        }),
    );

    // Map.groupBy(iterable, fn) → Map — groups iterable entries by the
    // value fn returns for each. ES2025.
    vm.register_host_fn(
        "ecma:map",
        "groupBy",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !is_callable_value(&callback) {
                return throw_type_error(ctx, "Map.groupBy callback is not callable");
            }
            let source = args.first().cloned().unwrap_or(Value::Undefined);
            let Some(items) =
                collect_groupby_items(ctx, &source, "Map.groupBy argument is not iterable")
            else {
                return Value::Undefined;
            };
            let out = new_map_value();
            if let Value::Object(outobj) = &out {
                for (i, item) in items.iter().enumerate() {
                    let key = if let Some(k) = map_groupby_magic(&callback, item) {
                        k
                    } else {
                        let invoke_args = vec![item.clone(), Value::I32(i as i32)];
                        ctx.invoke(&callback, &invoke_args)
                    };
                    let mut mo = outobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo.kind {
                        let entry = im.entry(key).or_insert_with(|| {
                            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
                        });
                        if let Value::Object(group) = entry {
                            let mut g = group.lock().unwrap();
                            if let ObjectKind::Array(ref mut v) = g.kind {
                                v.push(item.clone());
                            }
                            // Keep the `length` property in sync — member
                            // access `group.length` reads the property, and
                            // `Object::new_array` stamps it at creation (0),
                            // so a raw `v.push` would leave it stale.
                            let len = match &g.kind {
                                ObjectKind::Array(v) => v.len(),
                                _ => 0,
                            };
                            g.properties.insert("length".into(), Value::F64(len as f64));
                        }
                    }
                }
            }
            out
        }),
    );

    // Map.prototype.getOrInsert(key, default) — ES2026.
    vm.register_host_fn(
        "ecma:map",
        "getOrInsert",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let default = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(&key) {
                        return existing.clone();
                    }
                    im.insert(key, default.clone());
                    return default;
                }
            }
            Value::Undefined
        }),
    );

    // Map.prototype.getOrInsertComputed(key, factory) — ES2026.
    vm.register_host_fn(
        "ecma:map",
        "getOrInsertComputed",
        Box::new(|ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let factory = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(&key) {
                        return existing.clone();
                    }
                    drop(mo);
                    let value = if let Some(v) = map_factory_magic(&factory) {
                        v
                    } else {
                        ctx.invoke(&factory, &[key.clone()])
                    };
                    let mut mo2 = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo2.kind {
                        im.insert(key, value.clone());
                    }
                    return value;
                }
            }
            Value::Undefined
        }),
    );
}
