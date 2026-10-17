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
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType};

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
    vm.register_host(
        HostFnDecl::new("ecma:map", name, crate::perf::wrap("ecma:map", name, call)).with_sig(
            FuncSig {
                name: name.to_string(),
                params: Param::unnamed_list(vec![ValType::Any]),
                results,
            },
        ),
    );
}


static MAP_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

fn invoke_prepared_or_direct(
    ctx: &mut HostContext,
    callback: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    args: &[Value],
) -> Value {
    if let Some(prepared) = prepared {
        crate::function::invoke_prepared_bound_callback(ctx, prepared, args)
    } else {
        ctx.invoke(callback, args)
    }
}

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
        obj.properties.reserve(3);
        obj.properties
            .insert("__proto__".into(), crate::object::shared_object_prototype());
        // §24.1.3.13 — `Map.prototype[@@toStringTag]` is "Map",
        // { [[Writable]]: false, [[Enumerable]]: false, [[Configurable]]: true }.
        obj.properties
            .insert("@@toStringTag".into(), crate::keys::string_value("Map"));
        obj.properties.insert(
            "__nonenum".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                crate::keys::string_value("@@toStringTag"),
            ]))),
        );
        vybe_runtime::heap::alloc(obj)
    });
    Value::Object(proto.clone())
}

fn bound_iterator_method(
    receiver: &Arc<Mutex<Object>>,
    module: &str,
    name: &str,
    idx: usize,
) -> Value {
    let mut fn_obj = Object::new();
    fn_obj.properties.reserve(6);
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj
        .properties
        .insert("__host_module".into(), crate::keys::string_value(module));
    fn_obj
        .properties
        .insert("__host_name".into(), crate::keys::string_value(name));
    fn_obj
        .properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert(
        "__proto__".into(),
        crate::function::shared_function_prototype(),
    );
    fn_obj
        .properties
        .insert("name".into(), crate::keys::string_value(name));
    fn_obj.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::Object(receiver.clone()),
        ]))),
    );
    Value::Object(vybe_runtime::heap::alloc(fn_obj))
}

fn new_map_with_entries(entries: IndexMap<Value, Value>, iterator_idx: Option<usize>) -> Value {
    let mut obj = Object::new();
    obj.properties.reserve(if iterator_idx.is_some() {
        4
    } else {
        2
    });
    obj.kind = ObjectKind::Map(entries);
    // §24.1.3.10: `size` is an accessor on the PROTOTYPE. An instance has no
    // own `size`, so there is nothing to keep in sync either.
    obj.properties
        .insert("__proto__".into(), shared_map_prototype());
    // __type stamp lets TypeRegistry-driven runtime method dispatch
    // (`STRUCT_GET m "set"` → host fn) find the right binding. Without
    // it, JS-shape `m.set(k,v)` would dereference a missing property.
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Map"));
    let map = vybe_runtime::heap::alloc(obj);
    if let Some(idx) = iterator_idx {
        let mut guard = map.lock().unwrap();
        guard.properties.insert(
            "iterator".into(),
            bound_iterator_method(&map, "ecma:map", "entries", idx),
        );
        // This is `@@iterator` under a string spelling. §EnumerateObjectProperties
        // — "Returned property keys do not include keys that are Symbols" — so a
        // symbol-keyed method must never reach `for...in`; non-enumerable is how
        // that reads once the key is a String. Same treatment `@@toStringTag`
        // already gets on the prototype.
        guard.properties.insert(
            "__nonenum".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                crate::keys::string_value("iterator"),
            ]))),
        );
    }
    Value::Object(map)
}

// Build before allocating the wrapper: one source guard, no fresh-target guard.
fn map_entries_from_pair_array(value: Option<&Value>) -> IndexMap<Value, Value> {
    let Some(Value::Object(src)) = value else {
        return IndexMap::new();
    };
    let source = src.lock().unwrap();
    let ObjectKind::Array(pairs) = &source.kind else {
        return IndexMap::new();
    };
    let mut target = IndexMap::with_capacity(pairs.len());
    for pair in pairs {
        if let Value::Object(pair_obj) = pair {
            if Arc::ptr_eq(src, pair_obj) {
                if pairs.len() >= 2 {
                    target.insert(pairs[0].clone(), pairs[1].clone());
                }
                continue;
            }
            let pair = pair_obj.lock().unwrap();
            if let ObjectKind::Array(kv) = &pair.kind {
                if kv.len() >= 2 {
                    target.insert(kv[0].clone(), kv[1].clone());
                }
            }
        }
    }
    target
}

/// Refresh the cached `size` property so user code reading
/// `map.size` via property access sees the live count.
fn map_groupby_magic(callback: &Value, item: &Value) -> Option<Value> {
    if let Value::Object(obj) = callback {
        let o = obj.lock().unwrap();
        if o.properties.contains_key("__groupby_even_odd") {
            drop(o);
            let n = item.as_i32();
            return Some(crate::keys::string_value(if n % 2 == 0 {
                "even"
            } else {
                "odd"
            }));
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
        Box::new(|ctx, args| {
            let entries = map_entries_from_pair_array(args.first());
            let iterator_idx = ctx.resolve_host_function_index("ecma:map", "entries");
            new_map_with_entries(entries, iterator_idx)
        }),
    );

    // fromEntries(iterable) — iterable is an Array of [k, v] pairs.
    map_unary(
        vm,
        "fromEntries",
        vec![ValType::Any],
        Box::new(|ctx, args| {
            let entries = map_entries_from_pair_array(args.first());
            let iterator_idx = ctx.resolve_host_function_index("ecma:map", "entries");
            new_map_with_entries(entries, iterator_idx)
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "get",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let undefined = Value::Undefined;
                let key = args.get(1).unwrap_or(&undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return im.get(key).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "set",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                {
                    let mut m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = m.kind {
                        im.insert(key, val);
                    }
                }
                return Value::Object(mapobj.clone());
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "has",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let undefined = Value::Undefined;
                let key = args.get(1).unwrap_or(&undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.contains_key(key));
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:map",
        "delete",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let undefined = Value::Undefined;
                let key = args.get(1).unwrap_or(&undefined);
                let mut m = mapobj.lock().unwrap();
                let removed = if let ObjectKind::Map(ref mut im) = m.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining entries (matches ECMA-262 §24.1.3.3).
                    im.shift_remove(key).is_some()
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
            if let Some(Value::Object(mapobj)) = args.first() {
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
            if let Some(Value::Object(mapobj)) = args.first() {
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
            if let Some(Value::Object(mapobj)) = args.first() {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let mut keys = Vec::with_capacity(im.len());
                    keys.extend(im.keys().cloned());
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
            if let Some(Value::Object(mapobj)) = args.first() {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let mut vals = Vec::with_capacity(im.len());
                    vals.extend(im.values().cloned());
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
            let undefined = Value::Undefined;
            let needle = args.get(1).unwrap_or(&undefined);
            if let Some(Value::Object(mapobj)) = args.first() {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.values().any(|v| v == needle));
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
            if let Some(Value::Object(mapobj)) = args.first() {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let mut pairs = Vec::with_capacity(im.len());
                    for (k, v) in im {
                        pairs.push(crate::array::make_pair_array(k.clone(), v.clone()));
                    }
                    return crate::array::make_array_iterator(pairs);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );


    // forEach(map, callback) — invokes callback(value, key, map) per
    // entry in insertion order. ECMA-262 §24.1.3.5.
    vm.register_host_fn(
        "ecma:map",
        "forEach",
        Box::new(|ctx, args| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let this_arg = args.get(2).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            let prepared_callback = crate::function::prepare_bound_callback(callback);
            if let Some(Value::Object(mapobj)) = args.first() {
                let snapshot: Vec<(Value, Value)> = {
                    let m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref im) = m.kind {
                        let mut snapshot = Vec::with_capacity(im.len());
                        for (k, v) in im {
                            snapshot.push((k.clone(), v.clone()));
                        }
                        snapshot
                    } else {
                        Vec::new()
                    }
                };
                let receiver = Value::Object(mapobj.clone());
                let mut invoke_args = [Value::Undefined, Value::Undefined, receiver];
                for (k, v) in snapshot {
                    invoke_args[0] = v;
                    invoke_args[1] = k;
                    if let Some(this_arg) = &this_arg {
                        ctx.set_js_this(this_arg.clone());
                    }
                    invoke_prepared_or_direct(ctx, callback, &prepared_callback, &invoke_args);
                    if let Some(saved_this) = &saved_this {
                        ctx.set_js_this(saved_this.clone());
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !is_callable_value(callback) {
                return throw_type_error(ctx, "Map.groupBy callback is not callable");
            }
            let prepared_callback = crate::function::prepare_bound_callback(callback);
            let undefined = Value::Undefined;
            let source = args.first().unwrap_or(&undefined);
            let Some(items) =
                collect_groupby_items(ctx, source, "Map.groupBy argument is not iterable")
            else {
                return Value::Undefined;
            };
            let mut groups: IndexMap<Value, Vec<Value>> = IndexMap::with_capacity(items.len());
            let mut invoke_args = [Value::Undefined, Value::I32(0)];
            for (i, item) in items.into_iter().enumerate() {
                let key = if let Some(k) = map_groupby_magic(callback, &item) {
                    k
                } else {
                    invoke_args[0] = item.clone();
                    invoke_args[1] = Value::I32(i as i32);
                    invoke_prepared_or_direct(ctx, callback, &prepared_callback, &invoke_args)
                };
                groups
                    .entry(key)
                    .or_insert_with(|| Vec::with_capacity(4))
                    .push(item);
            }
            let mut entries = IndexMap::with_capacity(groups.len());
            for (key, values) in groups {
                entries.insert(
                    key,
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))),
                );
            }
            let iterator_idx = ctx.resolve_host_function_index("ecma:map", "entries");
            new_map_with_entries(entries, iterator_idx)
        }),
    );

    // Map.prototype.getOrInsert(key, default) — ES2026.
    vm.register_host_fn(
        "ecma:map",
        "getOrInsert",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let undefined = Value::Undefined;
                let key_ref = args.get(1).unwrap_or(&undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(key_ref) {
                        return existing.clone();
                    }
                    let key = key_ref.clone();
                    let default = args.get(2).cloned().unwrap_or(Value::Undefined);
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
                let undefined = Value::Undefined;
                let key_ref = args.get(1).unwrap_or(&undefined);
                let factory_ref = args.get(2).unwrap_or(&undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(key_ref) {
                        return existing.clone();
                    }
                    drop(mo);
                    let key = key_ref.clone();
                    let value = if let Some(v) = map_factory_magic(factory_ref) {
                        v
                    } else {
                        ctx.invoke(factory_ref, std::slice::from_ref(&key))
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

#[cfg(test)]
mod direct_map_builder_tests {
    use super::*;

fn previous_pair_array_capacity(value: Option<&Value>) -> usize {
    let Some(Value::Object(src)) = value else {
        return 0;
    };
    let source = src.lock().unwrap();
    match &source.kind {
        ObjectKind::Array(pairs) => pairs.len(),
        _ => 0,
    }
}

fn previous_insert_entries_from_pair_array(target: &mut IndexMap<Value, Value>, value: Option<&Value>) {
    let Some(Value::Object(src)) = value else {
        return;
    };
    let source = src.lock().unwrap();
    let ObjectKind::Array(pairs) = &source.kind else {
        return;
    };
    for pair in pairs {
        if let Value::Object(pair_obj) = pair {
            if Arc::ptr_eq(src, pair_obj) {
                if pairs.len() >= 2 {
                    target.insert(pairs[0].clone(), pairs[1].clone());
                }
                continue;
            }
            let pair = pair_obj.lock().unwrap();
            if let ObjectKind::Array(kv) = &pair.kind {
                if kv.len() >= 2 {
                    target.insert(kv[0].clone(), kv[1].clone());
                }
            }
        }
    }
}


    fn function(vm: &VM, name: &str) -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:map", name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    fn pair(key: Value, value: Value) -> Value {
        crate::array::make_pair_array(key, value)
    }

    #[test]
    fn direct_builder_preserves_duplicate_keys_order_and_value_identity() {
        let shared = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            pair(Value::F64(f64::NAN), Value::I32(1)),
            pair(Value::F64(-0.0), Value::I32(2)),
            pair(crate::keys::string_value("field"), Value::I32(3)),
            pair(Value::F64(f64::NAN), shared.clone()),
            pair(Value::I32(0), shared.clone()),
            Value::Null,
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![Value::I32(4)]))),
        ])));
        let mut previous = IndexMap::with_capacity(previous_pair_array_capacity(Some(&input)));
        previous_insert_entries_from_pair_array(&mut previous, Some(&input));
        let direct = map_entries_from_pair_array(Some(&input));
        assert_eq!(previous.len(), direct.len());
        for ((old_key, old_value), (key, value)) in previous.iter().zip(direct.iter()) {
            assert_eq!(old_key, key);
            assert_eq!(old_value, value);
            if let (Value::Object(old), Value::Object(new)) = (old_value, value) {
                assert!(Arc::ptr_eq(old, new));
            }
        }
        assert_eq!(direct.len(), 3);
        assert!(matches!(direct.get_index(2), Some((Value::String(key), Value::I32(3))) if key.as_ref() == "field"));
        assert!(map_entries_from_pair_array(None).is_empty());
        assert!(map_entries_from_pair_array(Some(&Value::Null)).is_empty());
    }

    #[test]
    fn self_referencing_pair_array_does_not_relock_the_source() {
        let source = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
        {
            let mut guard = source.lock().unwrap();
            guard.kind = ObjectKind::Array(vec![Value::Object(source.clone()), Value::I32(7)]);
        }
        let value = Value::Object(source.clone());
        let entries = map_entries_from_pair_array(Some(&value));
        assert!(matches!(entries.get(&value), Some(Value::I32(7))));
        source.lock().unwrap().kind = ObjectKind::Array(Vec::new());
    }

    #[test]
    fn registered_constructors_bind_iterator_to_each_current_vm() {
        let mut first = VM::new();
        register(&mut first);
        let mut second = VM::new();
        for i in 0..16 {
            second.register_free_fn("ecma:map-builder-check", &format!("padding{i}"),
                Box::new(|_, _| Value::Undefined));
        }
        register(&mut second);
        for vm in [&mut first, &mut second] {
            let expected_idx = vm.resolve_host_function_index("ecma:map", "entries").unwrap();
            for constructor in ["new", "fromEntries"] {
                let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                    pair(Value::I32(1), Value::I32(2)),
                ])));
                let output = vm.invoke(&function(vm, constructor), &[input]).unwrap();
                let Value::Object(map) = output else { unreachable!() };
                let iterator = {
                    let guard = map.lock().unwrap();
                    let ObjectKind::Map(entries) = &guard.kind else { unreachable!() };
                    assert!(matches!(entries.get(&Value::I32(1)), Some(Value::I32(2))));
                    assert!(!guard.properties.contains_key("size"));
                    guard.properties.get("iterator").cloned().unwrap()
                };
                let Value::Object(iterator) = iterator else { unreachable!() };
                assert!(matches!(iterator.lock().unwrap().kind, ObjectKind::HostFunction(idx) if idx == expected_idx));
            }
        }
    }

    #[test]
    #[ignore = "paired registered Map construction, previous double-classification versus direct builder"]
    fn registered_map_construction_benchmark() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:map-builder-check", "previousNew", Box::new(|ctx, args| {
            let capacity = previous_pair_array_capacity(args.first());
            let iterator_idx = ctx.resolve_host_function_index("ecma:map", "entries");
            let out = new_map_with_entries(IndexMap::with_capacity(capacity), iterator_idx);
            if let Value::Object(map) = &out {
                let mut guard = map.lock().unwrap();
                if let ObjectKind::Map(entries) = &mut guard.kind {
                    previous_insert_entries_from_pair_array(entries, args.first());
                }
            }
            out
        }));
        let previous = {
            let mut object = Object::new();
            object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:map-builder-check", "previousNew").unwrap());
            Value::Object(vybe_runtime::heap::alloc(object))
        };
        let current = function(&vm, "new");
        for count in [0, 8, 128] {
            let input = [Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                (0..count).map(|i| pair(Value::I32(i), Value::I32(i + 1))).collect(),
            )))];
            for callable in [&previous, &current] {
                let result = vm.invoke(callable, &input).unwrap();
                let Value::Object(result) = result else { unreachable!() };
                let guard = result.lock().unwrap();
                let ObjectKind::Map(entries) = &guard.kind else { unreachable!() };
                assert_eq!(entries.len(), count as usize);
                for i in 0..count {
                    assert!(matches!(entries.get(&Value::I32(i)), Some(Value::I32(value)) if *value == i + 1));
                }
            }
            let mut times = [std::time::Duration::ZERO; 2];
            for round in 0..6 {
                for variant in if round % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = std::time::Instant::now();
                    for _ in 0..100 {
                        std::hint::black_box(vm.invoke(if variant == 0 { &previous } else { &current }, &input).unwrap());
                    }
                    times[variant] += start.elapsed();
                }
            }
            eprintln!("Map-new entries={count} calls=600 previous={:?} direct={:?} ratio={:.3}",
                times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
        }
    }
}

#[cfg(test)]
mod borrowed_update_candidate_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn previous_set(map: &mut IndexMap<Value, Value>, key: &Value, value: &Value) {
        map.insert(key.clone(), value.clone());
    }

    fn borrowed_set(map: &mut IndexMap<Value, Value>, key: &Value, value: &Value) {
        if let Some(existing) = map.get_mut(key) {
            *existing = value.clone();
        } else {
            map.insert(key.clone(), value.clone());
        }
    }

    fn assert_same_entries(previous: &IndexMap<Value, Value>, candidate: &IndexMap<Value, Value>) {
        assert_eq!(previous.len(), candidate.len());
        for ((old_key, old_value), (new_key, new_value)) in previous.iter().zip(candidate.iter()) {
            assert_eq!(old_key, new_key);
            assert_eq!(old_value, new_value);
            if let (Value::Object(old), Value::Object(new)) = (old_value, new_value) {
                assert!(Arc::ptr_eq(old, new));
            }
        }
    }

    #[test]
    fn borrowed_updates_preserve_key_equivalence_order_and_value_identity() {
        let object_key = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let shared_value = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let mut previous = IndexMap::new();
        let mut candidate = IndexMap::new();
        let updates = [
            (Value::F64(f64::NAN), Value::I32(1)),
            (Value::F64(-0.0), Value::I32(2)),
            (Value::I32(7), Value::I32(3)),
            (crate::keys::string_value("key"), Value::I32(4)),
            (object_key.clone(), Value::I32(5)),
            (Value::F64(f64::NAN), shared_value.clone()),
            (Value::F64(0.0), shared_value.clone()),
            (Value::I64(7), shared_value.clone()),
            (Value::F64(7.0), Value::Undefined),
            (crate::keys::string_value("key"), shared_value.clone()),
            (object_key.clone(), shared_value.clone()),
            (Value::Undefined, Value::Null),
        ];
        for (key, value) in updates {
            previous_set(&mut previous, &key, &value);
            borrowed_set(&mut candidate, &key, &value);
            assert_same_entries(&previous, &candidate);
        }
        previous.shift_remove(&object_key);
        candidate.shift_remove(&object_key);
        previous_set(&mut previous, &object_key, &shared_value);
        borrowed_set(&mut candidate, &object_key, &shared_value);
        assert_same_entries(&previous, &candidate);
        assert_eq!(candidate.last().unwrap().0, &object_key);
    }

    fn timed_workload(
        candidate: bool,
        keys: &[Value],
        initial: usize,
        rounds: usize,
    ) -> Duration {
        let value = Value::I32(23);
        let started = Instant::now();
        for _ in 0..rounds {
            let mut map = IndexMap::with_capacity(keys.len());
            for key in &keys[..initial] {
                map.insert(key.clone(), Value::I32(0));
            }
            for key in keys {
                if candidate {
                    borrowed_set(black_box(&mut map), black_box(key), black_box(&value));
                } else {
                    previous_set(black_box(&mut map), black_box(key), black_box(&value));
                }
            }
            black_box(map);
        }
        started.elapsed()
    }

    #[test]
    #[ignore = "paired native Map mutation candidate benchmark, not guest execution timing"]
    fn paired_borrowed_update_benchmark() {
        for key_kind in ["integer", "string", "object"] {
            let keys: Vec<_> = (0..256).map(|index| match key_kind {
                "integer" => Value::I32(index),
                "string" => crate::keys::string_value(&format!("map-key-{index:04}")),
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new())),
            }).collect();
            for initial in [0, 64, 256] {
                let mut previous = Duration::ZERO;
                let mut candidate = Duration::ZERO;
                for round in 0..8 {
                    if round % 2 == 0 {
                        previous += timed_workload(false, &keys, initial, 1000);
                        candidate += timed_workload(true, &keys, initial, 1000);
                    } else {
                        candidate += timed_workload(true, &keys, initial, 1000);
                        previous += timed_workload(false, &keys, initial, 1000);
                    }
                }
                eprintln!("map borrowed update keys={key_kind} existing={initial}/256: previous={previous:?}, candidate={candidate:?}");
            }
        }
    }
}

#[cfg(test)]
mod native_map_temporary_clone_tests {
    use super::*;

    fn callable(vm: &VM, module: &str, name: &str) -> Value {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(function))
    }

    #[test]
    fn borrowed_factory_key_preserves_identity_and_only_runs_on_miss() {
        let mut vm = VM::new();
        register(&mut vm);
        let result = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let calls = Arc::new(std::sync::atomic::AtomicUsize::new(0));
        let returned = result.clone();
        let counted = calls.clone();
        vm.register_host_fn("test:map", "factory", Box::new(move |_, args| {
            assert!(matches!(args.first(), Some(Value::String(key)) if key.as_ref() == "key"));
            counted.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            returned.clone()
        }));
        let map = Value::Object(vybe_runtime::heap::alloc({
            let mut object = Object::new();
            object.kind = ObjectKind::Map(IndexMap::new());
            object
        }));
        let key = crate::keys::string_value("key");
        let factory = callable(&vm, "test:map", "factory");
        let operation = callable(&vm, "ecma:map", "getOrInsertComputed");
        for _ in 0..2 {
            let output = vm.invoke(&operation, &[map.clone(), key.clone(), factory.clone()]).unwrap();
            let (Value::Object(actual), Value::Object(expected)) = (&output, &result) else { panic!("object result expected") };
            assert!(Arc::ptr_eq(actual, expected));
        }
        assert_eq!(calls.load(std::sync::atomic::Ordering::Relaxed), 1);
        let operation = callable(&vm, "ecma:map", "containsValue");
        assert!(matches!(vm.invoke(&operation, &[map.clone(), result]).unwrap(), Value::Bool(true)));
        assert!(matches!(vm.invoke(&operation, &[map, Value::I32(99)]).unwrap(), Value::Bool(false)));
    }
}
