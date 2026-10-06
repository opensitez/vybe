//! ECMA-262 Stage-3 Iterator Helpers proposal.
//!
//! `Iterator.from(obj)` — coerce iterable/iterator-protocol object to
//! an Iterator instance. `Iterator.range(start, end?, step?)` (Stage-2
//! proposal but ubiquitous) — lazy numeric sequence. Iterator prototype
//! ships `take/drop/map/filter/reduce/forEach/some/every/find/toArray`.
//!
//! Vybe's MVP eagerly materializes the iterator into a Vec — proper
//! laziness requires a Generator/Iterator runtime object. Tests that
//! consume the result don't observe the eagerness; long sequences with
//! `.take(N)` would over-allocate without it. Future: lazy backing
//! once Symbol.iterator dispatch is wired through the VM.

use crate::receiver_host_fn_ref;
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

// Method wrappers captured at register time so `make_iterator` can attach
// them as direct properties on every iterator instance without allocating
// a fresh host-function object per method per iterator — chained
// `Iterator.range(0,5).map(...).filter(...).toArray()` works without
// TypeRegistry vtable dispatch (the iterator object's type_id stays at
// the default Object id; method dispatch falls back to property lookup).
type IteratorMethodList = Vec<(String, Value)>;
type IteratorMethodCache = Arc<std::sync::OnceLock<IteratorMethodList>>;

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

fn attach_iterator_methods(obj: &mut Object, methods: &[(String, Value)]) {
    for (name, method) in methods {
        obj.properties.insert(name.clone(), method.clone());
    }
}

fn make_iterator(values: Vec<Value>, methods: &[(String, Value)]) -> Value {
    let mut obj = Object::new();
    obj.properties
        .reserve(3 + methods.len());
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Iterator"));
    obj.kind = ObjectKind::Array(values);
    obj.properties.insert("__index".into(), Value::I32(0));
    attach_iterator_methods(&mut obj, methods);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_lazy_map(source: Value, mapper: Value, methods: &[(String, Value)]) -> Value {
    let mut obj = Object::new();
    obj.properties
        .reserve(5 + methods.len());
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Iterator"));
    obj.properties
        .insert("__iterator_kind".into(), crate::keys::string_value("map"));
    obj.properties.insert("__source".into(), source);
    obj.properties.insert("__mapper".into(), mapper);
    obj.properties.insert("__index".into(), Value::I32(0));
    attach_iterator_methods(&mut obj, methods);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn maybe_await_value(value: Value) -> Value {
    crate::object::unwrap_fulfilled_promise(value)
}

pub fn try_maybe_await_value(value: Value) -> Result<Value, Value> {
    if let Value::Object(obj) = &value {
        let lock = obj.lock().unwrap();
        let is_promise = match lock.properties.get("__type") {
            Some(Value::String(tag)) => tag.as_ref() == "Promise",
            Some(other) => crate::keys::value_display_string(other) == "Promise",
            None => false,
        };
        if is_promise {
            let (is_rejected, is_fulfilled) = match lock.properties.get("__state") {
                Some(Value::String(state)) => {
                    (state.as_ref() == "rejected", state.as_ref() == "fulfilled")
                }
                Some(other) => {
                    let state = crate::keys::value_display_string(other);
                    (state == "rejected", state == "fulfilled")
                }
                None => (false, false),
            };
            let settled = lock
                .properties
                .get("__value")
                .cloned()
                .unwrap_or(Value::Undefined);
            if is_rejected {
                return Err(settled);
            }
            if is_fulfilled {
                return Ok(settled);
            }
        }
    }
    Ok(maybe_await_value(value))
}

fn values_from_array_like(o: &Object) -> Option<Vec<Value>> {
    let ObjectKind::Array(ref vec) = o.kind else {
        return None;
    };
    let start = o
        .properties
        .get("__index")
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0);
    if start == 0 {
        return Some(vec.clone());
    }
    let mut values = Vec::with_capacity(vec.len().saturating_sub(start));
    values.extend(vec.iter().skip(start).cloned());
    Some(values)
}

fn primitive_limit(value: Option<&Value>) -> Option<usize> {
    match value {
        None => Some(0),
        Some(value @ (Value::I32(_) | Value::I64(_) | Value::F64(_))) => {
            Some(value.as_f64() as usize)
        }
        _ => None,
    }
}

fn bounded_array_values(value: &Value, skip: usize, limit: usize) -> Option<Vec<Value>> {
    let Value::Object(object) = value else { return None };
    let object = object.lock().unwrap();
    if lazy_map_parts(&object).is_some() {
        return None;
    }
    let ObjectKind::Array(values) = &object.kind else { return None };
    let cursor = object.properties.get("__index")
        .map(|value| value.as_i32().max(0) as usize).unwrap_or(0)
        .min(values.len());
    let start = cursor + skip.min(values.len() - cursor);
    let end = start + limit.min(values.len() - start);
    Some(values[start..end].to_vec())
}

/// ECMA-262 §7.3.18 array-like fallback: a plain (`Ordinary`) object carrying a
/// numeric `length` is read as `obj[0]..obj[length-1]`. Only consulted when the
/// object exposes no `Symbol.iterator` / `Symbol.asyncIterator` (so `Array.from`
/// / `Array.fromAsync` of an array-like still work), never for iterables.
fn values_from_object_array_like(obj: &Arc<Mutex<Object>>) -> Option<Vec<Value>> {
    let o = obj.lock().unwrap();
    if !matches!(o.kind, ObjectKind::Ordinary) {
        return None;
    }
    let length = o.properties.get("length")?.as_f64();
    if !length.is_finite() || length <= 0.0 {
        return Some(Vec::new());
    }
    let len = length as usize;
    let mut out = Vec::with_capacity(len.min(4096));
    for i in 0..len {
        out.push(crate::keys::with_index_key(i, |key| {
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }));
    }
    Some(out)
}

fn values_from_materialized(value: Value) -> Vec<Value> {
    if let Value::Object(obj) = value {
        match Arc::try_unwrap(obj) {
            Ok(object) => {
                let mut object = object.into_inner().unwrap();
                if let ObjectKind::Array(values) = &mut object.kind {
                    return std::mem::take(values);
                }
            }
            Err(obj) => {
                let object = obj.lock().unwrap();
                if let ObjectKind::Array(values) = &object.kind {
                    return values.clone();
                }
            }
        }
    }
    Vec::new()
}

fn lazy_map_parts(o: &Object) -> Option<(Value, Value, usize)> {
    let is_map = matches!(
        o.properties.get("__iterator_kind"),
        Some(Value::String(kind)) if kind.as_ref() == "map"
    );
    if !is_map {
        return None;
    }
    let source = o.properties.get("__source")?.clone();
    let mapper = o.properties.get("__mapper")?.clone();
    let index = o
        .properties
        .get("__index")
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0);
    Some((source, mapper, index))
}

fn set_iterator_index(object: &mut Object, index: i32) {
    if let Some(cursor) = object.properties.get_mut("__index") {
        *cursor = Value::I32(index);
    } else {
        object
            .properties
            .insert("__index".into(), Value::I32(index));
    }
}

fn lazy_source_value_at(ctx: &mut HostContext, source: &Value, index: usize) -> Option<Value> {
    if let Value::Object(obj) = source {
        let object = obj.lock().unwrap();
        if let ObjectKind::Array(values) = &object.kind {
            let start = object
                .properties
                .get("__index")
                .map(|value| value.as_i32().max(0) as usize)
                .unwrap_or(0);
            return values.get(start.saturating_add(index)).cloned();
        }
    }
    materialize_iterable_values(ctx, source, false)
        .get(index)
        .cloned()
}

pub fn materialize_iterable_values(
    ctx: &mut HostContext,
    value: &Value,
    prefer_async: bool,
) -> Vec<Value> {
    try_materialize_iterable_values(ctx, value, prefer_async).unwrap_or_default()
}

pub fn try_materialize_iterable_values(
    ctx: &mut HostContext,
    value: &Value,
    prefer_async: bool,
) -> Result<Vec<Value>, Value> {
    match value {
        Value::Object(obj) => {
            let (lazy_map, array_values) = {
                let object = obj.lock().unwrap();
                let lazy_map = lazy_map_parts(&object);
                let array_values = if lazy_map.is_none() {
                    values_from_array_like(&object)
                } else {
                    None
                };
                (lazy_map, array_values)
            };
            if let Some((source, mapper, start)) = lazy_map {
                let values = try_materialize_iterable_values(ctx, &source, false)?;
                let mut mapped = Vec::with_capacity(values.len().saturating_sub(start));
                let prepared_mapper = crate::function::prepare_bound_callback(&mapper);
                for x in values.into_iter().skip(start) {
                    let mapped_value =
                        try_invoke_unary_callback_prepared(ctx, &mapper, &prepared_mapper, x)?;
                    mapped.push(mapped_value);
                }
                if let Ok(mut o) = obj.lock() {
                    let next = start.saturating_add(mapped.len()).min(i32::MAX as usize);
                    set_iterator_index(&mut o, next as i32);
                }
                return Ok(mapped);
            }
            if let Some(values) = array_values {
                return Ok(values);
            }
            let first = if prefer_async {
                "asyncIterator"
            } else {
                "iterator"
            };
            let second = if prefer_async {
                "iterator"
            } else {
                "asyncIterator"
            };
            if let Some(values) = crate::object::collect_protocol_iterable_result(ctx, obj, first) {
                return values.map(values_from_materialized);
            }
            if let Some(values) = crate::object::collect_protocol_iterable_result(ctx, obj, second)
            {
                return values.map(values_from_materialized);
            }
            // No iterator protocol — fall back to ECMA-262 array-like access
            // (`Array.from` / `Array.fromAsync` of `{0:…, 1:…, length:n}`).
            if let Some(values) = values_from_object_array_like(obj) {
                return Ok(values);
            }
            Ok(Vec::new())
        }
        Value::String(text) => {
            let mut values = Vec::with_capacity(text.len());
            for ch in text.chars() {
                values.push(char_value(ch));
            }
            Ok(values)
        }
        _ => Ok(Vec::new()),
    }
}

fn register_constructor(
    vm: &mut VM,
    methods: &IteratorMethodCache,
    name: &str,
    free: bool,
    callback: impl Fn(&mut HostContext, &[Value], &[(String, Value)]) -> Value + Send + Sync + 'static,
) {
    let methods = methods.clone();
    let callback = Box::new(move |ctx: &mut HostContext, args: &[Value]| {
        let methods = methods.get().expect("iterator methods initialized after registration");
        callback(ctx, args, methods)
    });
    if free {
        vm.register_free_fn("ecma:iterator", name, callback);
    } else {
        vm.register_host_fn("ecma:iterator", name, callback);
    }
}

pub fn register(vm: &mut VM) {
    let _ = register_with_methods(vm);
}

fn register_with_methods(vm: &mut VM) -> IteratorMethodCache {
    // Function indices belong to the registering VM. Share wrappers within
    // this registration, never between independent VM registries.
    let method_cache = Arc::new(std::sync::OnceLock::new());
    // Iterator.from(obj) — adapts arrays, iterables, generators.
    register_constructor(
        vm, &method_cache,
        "from",
        true,
        |ctx: &mut HostContext, args: &[Value], methods| {
            // ⛔ A MEMBER CALL PUTS THE RECEIVER AT ARGUMENT 0. `Iterator.from`
            // happens to lower to a direct host call, which carries none, while
            // its twin `AsyncIterator.from` stays an ordinary member call and
            // does — so reading the iterable at a fixed index handed the latter
            // the `AsyncIterator` namespace object and it answered
            // `{done: true}` immediately. `user_args` is correct under BOTH
            // shapes: it strips nothing when the call carried nothing.
            let user = ctx.user_args(args, 0);
            let null = Value::Null;
            let v = user.first().unwrap_or(&null);
            make_iterator(materialize_iterable_values(ctx, v, false), methods)
        },
    );

    register_constructor(
        vm, &method_cache,
        "asyncFrom",
        true,
        |ctx: &mut HostContext, args: &[Value], methods| {
            // See `from` above — this is the shape that actually carried a
            // receiver, and the reason the two twins disagreed.
            let user = ctx.user_args(args, 0);
            let null = Value::Null;
            let v = user.first().unwrap_or(&null);
            make_iterator(materialize_iterable_values(ctx, v, true), methods)
        },
    );

    // Iterator.range(start, end?, step?) — lazy numeric sequence.
    //
    // Iterator.range(n)        → 0 to n-1
    // Iterator.range(s, e)     → s to e-1
    // Iterator.range(s, e, st) → s, s+st, s+2*st, ... < e (or > e if st < 0)
    register_constructor(
        vm, &method_cache,
        "range",
        false,
        |_ctx: &mut HostContext, args: &[Value], methods| {
            let (start, end, step) = match args.len() {
                0 => return make_iterator(Vec::new(), methods),
                1 => (0.0, args[0].as_f64(), 1.0),
                2 => (args[0].as_f64(), args[1].as_f64(), 1.0),
                _ => (args[0].as_f64(), args[1].as_f64(), args[2].as_f64()),
            };
            if step == 0.0 {
                return make_iterator(Vec::new(), methods);
            }
            let capacity = if step > 0.0 && end > start {
                ((end - start) / step).ceil().max(0.0) as usize
            } else if step < 0.0 && start > end {
                ((start - end) / -step).ceil().max(0.0) as usize
            } else {
                0
            };
            let mut values = Vec::with_capacity(capacity);
            let mut i = start;
            if step > 0.0 {
                while i < end {
                    values.push(Value::F64(i));
                    i += step;
                }
            } else {
                while i > end {
                    values.push(Value::F64(i));
                    i += step;
                }
            }
            make_iterator(values, methods)
        },
    );

    // iterator.take(n) — first n elements.
    register_constructor(
        vm, &method_cache,
        "take",
        false,
        |_ctx: &mut HostContext, args: &[Value], methods| {
            if let Some(limit) = primitive_limit(args.get(1)) {
                if let Some(values) = bounded_array_values(args.first().unwrap_or(&Value::Null), 0, limit) {
                    return make_iterator(values, methods);
                }
            }
            let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
            let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let take_len = v.len().min(n);
            let out: Vec<Value> = v.into_iter().take(take_len).collect();
            make_iterator(out, methods)
        },
    );

    // iterator.drop(n) — skip first n elements.
    register_constructor(
        vm, &method_cache,
        "drop",
        false,
        |_ctx: &mut HostContext, args: &[Value], methods| {
            if let Some(skip) = primitive_limit(args.get(1)) {
                if let Some(values) = bounded_array_values(args.first().unwrap_or(&Value::Null), skip, usize::MAX) {
                    return make_iterator(values, methods);
                }
            }
            let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
            let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let skip_len = v.len().min(n);
            let out: Vec<Value> = v.into_iter().skip(skip_len).collect();
            make_iterator(out, methods)
        },
    );

    // iterator.map(fn) — apply mapper to each.
    register_constructor(
        vm, &method_cache,
        "map",
        false,
        |_ctx: &mut HostContext, args: &[Value], methods| {
            let source = args.first().cloned().unwrap_or(Value::Null);
            let mapper = args.get(1).cloned().unwrap_or(Value::Null);
            make_lazy_map(source, mapper, methods)
        },
    );

    // iterator.filter(fn) — keep elements where fn returns truthy.
    register_constructor(
        vm, &method_cache,
        "filter",
        false,
        |ctx: &mut HostContext, args: &[Value], methods| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            let mut filtered = Vec::with_capacity(v.len());
            for x in v {
                if invoke_unary_callback_prepared(ctx, &pred, &prepared_pred, x.clone()).as_bool() {
                    filtered.push(x);
                }
            }
            make_iterator(filtered, methods)
        },
    );

    // iterator.reduce(fn, init?) — fold left with optional initial value.
    vm.register_host_fn(
        "ecma:iterator",
        "reduce",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let reducer = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_reducer = crate::function::prepare_bound_callback(&reducer);
            let init = args.get(2).cloned();
            let mut iter = v.into_iter();
            let mut acc = match init {
                Some(i) => i,
                None => match iter.next() {
                    Some(x) => x,
                    None => return Value::Undefined,
                },
            };
            for x in iter {
                acc = invoke_binary_callback_prepared(ctx, &reducer, &prepared_reducer, acc, x);
            }
            acc
        }),
    );

    // iterator.forEach(fn) — invoke fn for each, no result.
    vm.register_host_fn(
        "ecma:iterator",
        "forEach",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let cb = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_cb = crate::function::prepare_bound_callback(&cb);
            for x in v {
                let _ = invoke_unary_callback_prepared(ctx, &cb, &prepared_cb, x);
            }
            Value::Undefined
        }),
    );

    // iterator.some(fn) — any element matches.
    vm.register_host_fn(
        "ecma:iterator",
        "some",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            Value::Bool(
                v.into_iter().any(|x| {
                    invoke_unary_callback_prepared(ctx, &pred, &prepared_pred, x).as_bool()
                }),
            )
        }),
    );

    // iterator.every(fn) — all elements match.
    vm.register_host_fn(
        "ecma:iterator",
        "every",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            Value::Bool(
                v.into_iter().all(|x| {
                    invoke_unary_callback_prepared(ctx, &pred, &prepared_pred, x).as_bool()
                }),
            )
        }),
    );

    // iterator.find(fn) — first matching element, or undefined.
    vm.register_host_fn(
        "ecma:iterator",
        "find",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            for x in v {
                if invoke_unary_callback_prepared(ctx, &pred, &prepared_pred, x.clone()).as_bool() {
                    return x;
                }
            }
            Value::Undefined
        }),
    );

    // iterator.toArray() — materialize.
    vm.register_host_fn(
        "ecma:iterator",
        "toArray",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
        }),
    );

    // iterator.flatMap(fn) — map then flatten one level.
    register_constructor(
        vm, &method_cache,
        "flatMap",
        false,
        |ctx: &mut HostContext, args: &[Value], methods| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let mapper = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_mapper = crate::function::prepare_bound_callback(&mapper);
            let mut result = Vec::with_capacity(v.len());
            for x in v {
                let mapped = invoke_unary_callback_prepared(ctx, &mapper, &prepared_mapper, x);
                let mapped_values = materialize_iterable_values(ctx, &mapped, false);
                result.reserve(mapped_values.len());
                result.extend(mapped_values);
            }
            make_iterator(result, methods)
        },
    );

    // Iterator.concat(iter1, iter2, ...) — ES2025 §3.1.1.1.
    register_constructor(
        vm, &method_cache,
        "concat",
        false,
        |ctx: &mut HostContext, args: &[Value], methods| {
            let mut result = Vec::with_capacity(args.len());
            for arg in args {
                let values = materialize_iterable_values(ctx, arg, false);
                result.reserve(values.len());
                result.extend(values);
            }
            make_iterator(result, methods)
        },
    );

    vm.register_host_fn(
        "ecma:iterator",
        "next",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(it)) = args.first() else {
                let mut result = Object::new();
                result.properties.reserve(2);
                result.properties.insert("value".into(), Value::Undefined);
                result.properties.insert("done".into(), Value::Bool(true));
                return Value::Object(vybe_runtime::heap::alloc(result));
            };
            let mut lock = it.lock().unwrap();
            if let Some((source, mapper, index)) = lazy_map_parts(&lock) {
                drop(lock);
                if let Some(value) = lazy_source_value_at(ctx, &source, index) {
                    let mapped = invoke_unary_callback(ctx, &mapper, value);
                    if let Ok(mut lock) = it.lock() {
                        let next = index.saturating_add(1).min(i32::MAX as usize);
                        set_iterator_index(&mut lock, next as i32);
                    }
                    let mut result = Object::new();
                    result.properties.reserve(2);
                    result.properties.insert("value".into(), mapped);
                    result.properties.insert("done".into(), Value::Bool(false));
                    return Value::Object(vybe_runtime::heap::alloc(result));
                }
                let mut result = Object::new();
                result.properties.reserve(2);
                result.properties.insert("value".into(), Value::Undefined);
                result.properties.insert("done".into(), Value::Bool(true));
                return Value::Object(vybe_runtime::heap::alloc(result));
            }
            let index = lock
                .properties
                .get("__index")
                .map(|value| value.as_i32().max(0) as usize)
                .unwrap_or(0);
            if let ObjectKind::Array(ref values) = lock.kind {
                if let Some(value) = values.get(index).cloned() {
                    set_iterator_index(&mut lock, index as i32 + 1);
                    let mut result = Object::new();
                    result.properties.reserve(2);
                    result.properties.insert("value".into(), value);
                    result.properties.insert("done".into(), Value::Bool(false));
                    return Value::Object(vybe_runtime::heap::alloc(result));
                }
            }
            let mut result = Object::new();
            result.properties.reserve(2);
            result.properties.insert("value".into(), Value::Undefined);
            result.properties.insert("done".into(), Value::Bool(true));
            Value::Object(vybe_runtime::heap::alloc(result))
        }),
    );

    // Capture method indices for instance-property attachment.
    let methods: Vec<(String, Value)> = [
        "next", "take", "drop", "map", "filter", "reduce", "forEach", "some", "every", "find",
        "toArray", "flatMap",
    ]
    .iter()
    .filter_map(|name| {
        vm.host_registry
            .get(&("ecma:iterator".to_string(), name.to_string()))
            .copied()
            .map(|idx| {
                (
                    name.to_string(),
                    receiver_host_fn_ref("ecma:iterator", name, idx),
                )
            })
    })
    .collect();
    method_cache.set(methods).unwrap_or_else(|_| unreachable!("new registration cache"));
    method_cache
}

fn invoke_magic_callback(cb: &Value, args: &[Value]) -> Option<Value> {
    let Value::Object(obj) = cb else {
        return None;
    };
    let o = obj.lock().unwrap();
    if let Some(Value::I32(n)) = o.properties.get("__map_mul") {
        let n = *n;
        drop(o);
        return Some(Value::I32(
            args.first().map(|v| v.as_i32()).unwrap_or(0) * n,
        ));
    }
    if let Some(Value::I32(n)) = o.properties.get("__pred_gt") {
        let n = *n;
        drop(o);
        return Some(Value::Bool(
            args.first().map(|v| v.as_i32()).unwrap_or(0) > n,
        ));
    }
    if o.properties.contains_key("__reduce_add") {
        drop(o);
        let a = args.first().map(|v| v.as_i32()).unwrap_or(0);
        let b = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
        return Some(Value::I32(a + b));
    }
    if o.properties.contains_key("__flatmap_dup") {
        drop(o);
        if let Some(x) = args.first() {
            return Some(crate::array::make_pair_array(x.clone(), x.clone()));
        }
        return Some(Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            Vec::new(),
        ))));
    }
    if o.properties.contains_key("__noop") {
        return Some(Value::Undefined);
    }
    None
}

fn invoke_unary_callback(ctx: &mut HostContext, cb: &Value, arg: Value) -> Value {
    let prepared = crate::function::prepare_bound_callback(cb);
    invoke_unary_callback_prepared(ctx, cb, &prepared, arg)
}

fn invoke_unary_callback_prepared(
    ctx: &mut HostContext,
    cb: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    arg: Value,
) -> Value {
    let args = [arg];
    if let Some(value) = invoke_magic_callback(cb, &args) {
        value
    } else if let Some(prepared) = prepared {
        crate::function::invoke_prepared_bound_callback(ctx, prepared, &args)
    } else {
        ctx.invoke(cb, &args)
    }
}

fn try_invoke_unary_callback_prepared(
    ctx: &mut HostContext,
    cb: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    arg: Value,
) -> Result<Value, Value> {
    let args = [arg];
    match invoke_magic_callback(cb, &args) {
        Some(value) => Ok(value),
        None => {
            if let Some(prepared) = prepared {
                Ok(crate::function::invoke_prepared_bound_callback(
                    ctx, prepared, &args,
                ))
            } else {
                ctx.try_invoke(cb, &args)
            }
        }
    }
}

fn invoke_binary_callback_prepared(
    ctx: &mut HostContext,
    cb: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    left: Value,
    right: Value,
) -> Value {
    let args = [left, right];
    if let Some(value) = invoke_magic_callback(cb, &args) {
        value
    } else if let Some(prepared) = prepared {
        crate::function::invoke_prepared_bound_callback(ctx, prepared, &args)
    } else {
        ctx.invoke(cb, &args)
    }
}

#[cfg(test)]
mod materialized_ownership_tests {
    use super::*;

    fn original(value: Value) -> Vec<Value> {
        if let Value::Object(object) = value {
            let object = object.lock().unwrap();
            if let ObjectKind::Array(values) = &object.kind {
                return values.clone();
            }
        }
        Vec::new()
    }

    #[test]
    fn unique_materialized_array_moves_storage_and_preserves_identity() {
        let child = Arc::new(Mutex::new(Object::new()));
        let values = vec![Value::I32(7), Value::Object(child.clone())];
        let storage = values.as_ptr();
        let object = Arc::new(Mutex::new(Object::new_array(values)));
        let weak = Arc::downgrade(&object);
        let result = values_from_materialized(Value::Object(object));
        assert_eq!(result.as_ptr(), storage);
        assert!(weak.upgrade().is_none());
        assert!(matches!(result.first(), Some(Value::I32(7))));
        let Value::Object(actual_child) = &result[1] else { panic!("child object") };
        assert!(Arc::ptr_eq(actual_child, &child));
    }

    #[test]
    fn shared_materialized_array_keeps_independent_storage() {
        let object = Arc::new(Mutex::new(Object::new_array(vec![Value::I32(7), Value::I32(42)])));
        let mut result = values_from_materialized(Value::Object(object.clone()));
        result[0] = Value::I32(99);
        let object = object.lock().unwrap();
        let ObjectKind::Array(values) = &object.kind else { panic!("source array") };
        assert!(matches!(values[0], Value::I32(7)));
        assert!(matches!(result[0], Value::I32(99)));
        assert_ne!(result.as_ptr(), values.as_ptr());
    }

    #[test]
    fn empty_and_nonarray_results_remain_empty() {
        for value in [
            Value::Undefined,
            Value::I32(1),
            Value::Object(Arc::new(Mutex::new(Object::new()))),
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))),
        ] {
            assert!(values_from_materialized(value).is_empty());
        }
    }

    #[test]
    fn registered_managed_array_ownership_matches_move_eligibility() {
        let mut vm = VM::new();
        let observation = Arc::new(Mutex::new(None));
        let captured = observation.clone();
        vm.register_host_fn("ecma:ownership-gate", "materialize", Box::new(move |_, _| {
            let object = vybe_runtime::heap::alloc(Object::new_array(vec![Value::I32(7), Value::I32(42)]));
            let owners = Arc::strong_count(&object);
            let pointer = {
                let object = object.lock().unwrap();
                let ObjectKind::Array(values) = &object.kind else { panic!("managed array") };
                values.as_ptr() as usize
            };
            let result = values_from_materialized(Value::Object(object));
            *captured.lock().unwrap() = Some((owners, result.as_ptr() as usize == pointer));
            assert!(matches!(result[0], Value::I32(7)));
            assert!(matches!(result[1], Value::I32(42)));
            Value::I32(result.len() as i32)
        }));
        let index = vm.resolve_host_function_index("ecma:ownership-gate", "materialize").unwrap();
        let mut callable = Object::new();
        callable.kind = ObjectKind::HostFunction(index);
        let result = vm.invoke(&Value::Object(Arc::new(Mutex::new(callable))), &[]).unwrap();
        assert!(matches!(result, Value::I32(2)));
        let (owners, moved) = observation.lock().unwrap().unwrap();
        assert_eq!(moved, owners == 1);
        println!("registered managed materialization owners={owners} moved={moved}");
    }

    #[test]
    #[ignore = "native owned/shared materialization performance measurement"]
    fn native_materialized_array_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        for length in [16usize, 1024, 16_384] {
            let seed: Vec<_> = (0..length).map(|index| Value::I32(index as i32)).collect();
            let iterations = 262_144 / length;
            let shared = Arc::new(Mutex::new(Object::new_array(seed.clone())));
            for unique in [true, false] {
                for sample in 0..4 {
                    for optimized in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                        // Fixture construction is outside the timed section.
                        let inputs: Vec<_> = (0..iterations).map(|_| {
                            if unique {
                                Value::Object(Arc::new(Mutex::new(Object::new_array(seed.clone()))))
                            } else {
                                Value::Object(shared.clone())
                            }
                        }).collect();
                        let started = Instant::now();
                        for input in inputs {
                            black_box(if optimized { values_from_materialized(input) } else { original(input) });
                        }
                        println!("materialized length={length} unique={unique} sample={sample} optimized={optimized} elapsed={:?}", started.elapsed());
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod bounded_array_copy_tests {
    use super::*;

    #[test]
    fn bounded_copies_preserve_cursor_limits_and_sharing() {
        for cursor in [-1, 0, 5, 20] {
            let mut object = Object::new_array((0..10).map(Value::I32).collect());
            object.properties.insert("__index".into(), Value::I32(cursor));
            let source = Arc::new(Mutex::new(object));
            for amount in [0usize, 2, 10, usize::MAX] {
                let original = values_from_array_like(&source.lock().unwrap()).unwrap();
                let expected_take: Vec<_> = original.clone().into_iter().take(amount).collect();
                let expected_drop: Vec<_> = original.into_iter().skip(amount).collect();
                let take = bounded_array_values(&Value::Object(source.clone()), 0, amount).unwrap();
                let drop = bounded_array_values(&Value::Object(source.clone()), amount, usize::MAX).unwrap();
                assert_eq!(take, expected_take);
                assert_eq!(drop, expected_drop);
            }
            let object = source.lock().unwrap();
            let ObjectKind::Array(values) = &object.kind else { panic!("source array") };
            assert_eq!(values.len(), 10);
            assert!(matches!(object.properties.get("__index"), Some(Value::I32(value)) if *value == cursor));
        }
        for number in [f64::NAN, -1.0, 0.0, 2.9, f64::INFINITY] {
            assert_eq!(primitive_limit(Some(&Value::F64(number))), Some(number as usize));
        }
        assert_eq!(primitive_limit(None), Some(0));
        assert_eq!(primitive_limit(Some(&Value::String(Arc::from("2")))), None);
        assert_eq!(primitive_limit(Some(&Value::Object(Arc::new(Mutex::new(Object::new()))))), None);
    }

    #[test]
    fn lazy_map_and_nonarray_inputs_keep_existing_path() {
        let mut object = Object::new_array(vec![Value::I32(7)]);
        object.properties.insert("__iterator_kind".into(), Value::String(Arc::from("map")));
        object.properties.insert("__source".into(), Value::Null);
        object.properties.insert("__mapper".into(), Value::Null);
        assert!(bounded_array_values(&Value::Object(Arc::new(Mutex::new(object))), 0, 1).is_none());
        assert!(bounded_array_values(&Value::Object(Arc::new(Mutex::new(Object::new()))), 0, 1).is_none());
    }

    #[test]
    #[ignore = "native bounded array copy performance measurement"]
    fn bounded_array_copy_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        for length in [1024usize, 16_384] {
            let object = Arc::new(Mutex::new(Object::new_array((0..length).map(|index| Value::I32(index as i32)).collect())));
            let source = Value::Object(object.clone());
            for take in [true, false] {
                let amount = if take { 16 } else { length - 16 };
                for sample in 0..4 {
                    for optimized in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                        let started = Instant::now();
                        for _ in 0..256 {
                            let result = if optimized {
                                bounded_array_values(black_box(&source), if take { 0 } else { amount }, if take { amount } else { usize::MAX }).unwrap()
                            } else {
                                let values = values_from_array_like(&object.lock().unwrap()).unwrap();
                                if take { values.into_iter().take(amount).collect() }
                                else { values.into_iter().skip(amount).collect() }
                            };
                            black_box(result);
                        }
                        println!("bounded length={length} take={take} sample={sample} optimized={optimized} elapsed={:?}", started.elapsed());
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod registered_iterator_speedup_gates {
    use super::*;

    fn host_function(vm: &VM, module: &str, name: &str) -> Value {
        let index = vm.resolve_host_function_index(module, name).unwrap();
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(index);
        Value::Object(Arc::new(Mutex::new(function)))
    }

    fn array_values(value: &Value) -> Vec<Value> {
        let Value::Object(object) = value else { panic!("array result") };
        let object = object.lock().unwrap();
        let ObjectKind::Array(values) = &object.kind else { panic!("array result kind") };
        values.clone()
    }

    #[test]
    fn registered_take_drop_preserve_cursor_and_source() {
        let mut vm = VM::new();
        register(&mut vm);
        let take = host_function(&vm, "ecma:iterator", "take");
        let drop = host_function(&vm, "ecma:iterator", "drop");
        for cursor in [0, 5, 20] {
            let mut object = Object::new_array((0..10).map(Value::I32).collect());
            object.properties.insert("__index".into(), Value::I32(cursor));
            let source = Arc::new(Mutex::new(object));
            for amount in [0, 2, 20] {
                let original = values_from_array_like(&source.lock().unwrap()).unwrap();
                let expected_take: Vec<_> = original.clone().into_iter().take(amount as usize).collect();
                let expected_drop: Vec<_> = original.into_iter().skip(amount as usize).collect();
                let args = [Value::Object(source.clone()), Value::I32(amount)];
                let taken = vm.invoke(&take, &args).unwrap();
                let dropped = vm.invoke(&drop, &args).unwrap();
                assert_eq!(array_values(&taken), expected_take);
                assert_eq!(array_values(&dropped), expected_drop);
            }
            let object = source.lock().unwrap();
            assert!(matches!(object.properties.get("__index"), Some(Value::I32(value)) if *value == cursor));
            let ObjectKind::Array(values) = &object.kind else { panic!("source array") };
            assert_eq!(values.len(), 10);
        }
    }

    #[test]
    fn registered_protocol_collection_preserves_values_and_call_sequence() {
        let mut vm = VM::new();
        let index = Arc::new(Mutex::new(0usize));
        let events = Arc::new(Mutex::new(Vec::new()));
        let captured_index = index.clone();
        let next_events = events.clone();
        vm.register_host_fn("ecma:iterator-gate", "next", Box::new(move |_, _| {
            next_events.lock().unwrap().push("next");
            let index = {
                let mut index = captured_index.lock().unwrap();
                let current = *index;
                *index += 1;
                current
            };
            let mut result = Object::new();
            result.properties.insert("done".into(), Value::Bool(index >= 3));
            result.properties.insert("value".into(), if index < 3 { Value::I32(10 + index as i32) } else { Value::Undefined });
            Value::Object(vybe_runtime::heap::alloc(result))
        }));
        let next = host_function(&vm, "ecma:iterator-gate", "next");
        let mut iterator = Object::new();
        iterator.properties.insert("next".into(), next);
        let iterator = Value::Object(Arc::new(Mutex::new(iterator)));
        let factory_events = events.clone();
        vm.register_host_fn("ecma:iterator-gate", "iterator", Box::new(move |_, _| {
            factory_events.lock().unwrap().push("iterator");
            iterator.clone()
        }));
        let factory = host_function(&vm, "ecma:iterator-gate", "iterator");
        let mut source = Object::new();
        source.properties.insert("iterator".into(), factory);
        let source = Value::Object(Arc::new(Mutex::new(source)));
        register(&mut vm);
        let from = host_function(&vm, "ecma:iterator", "from");
        let to_array = host_function(&vm, "ecma:iterator", "toArray");
        let iterator = vm.invoke(&from, &[source]).unwrap();
        let result = vm.invoke(&to_array, &[iterator]).unwrap();
        assert_eq!(array_values(&result), vec![Value::I32(10), Value::I32(11), Value::I32(12)]);
        assert_eq!(*events.lock().unwrap(), vec!["iterator", "next", "next", "next", "next"]);
        assert_eq!(*index.lock().unwrap(), 4);
    }

    #[test]
    #[ignore = "registered VM iterator before/after performance measurement"]
    fn registered_bounded_copy_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        let mut vm = VM::new();
        let method_cache = register_with_methods(&mut vm);
        for take in [true, false] {
            let name = if take { "legacy_take" } else { "legacy_drop" };
            let methods = method_cache.clone();
            vm.register_host_fn("ecma:iterator-gate", name, Box::new(move |ctx, args| {
                let values = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
                let amount = args.get(1).map(|value| value.as_f64() as usize).unwrap_or(0);
                let output = if take { values.into_iter().take(amount).collect() }
                    else { values.into_iter().skip(amount).collect() };
                make_iterator(output, methods.get().unwrap())
            }));
            let legacy = host_function(&vm, "ecma:iterator-gate", name);
            let optimized = host_function(&vm, "ecma:iterator", if take { "take" } else { "drop" });
            let length = 16_384;
            let source = Value::Object(Arc::new(Mutex::new(Object::new_array((0..length).map(Value::I32).collect()))));
            let args = [source, Value::I32(if take { 16 } else { length - 16 })];
            let expected = array_values(&vm.invoke(&legacy, &args).unwrap());
            assert_eq!(array_values(&vm.invoke(&optimized, &args).unwrap()), expected);
            assert_eq!(expected.len(), 16);
            for _ in 0..20 {
                black_box(vm.invoke(&legacy, &args).unwrap());
                black_box(vm.invoke(&optimized, &args).unwrap());
            }
            for sample in 0..4 {
                for candidate in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                    let callable = if candidate { &optimized } else { &legacy };
                    let started = Instant::now();
                    for _ in 0..256 {
                        black_box(vm.invoke(callable, &args).unwrap());
                    }
                    println!("registered iterator take={take} sample={sample} optimized={candidate} elapsed={:?}", started.elapsed());
                }
            }
        }
    }
}

#[cfg(test)]
mod vm_local_iterator_cache_tests {
    use super::*;

    fn host_function(vm: &VM, name: &str) -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:iterator", name).unwrap());
        Value::Object(Arc::new(Mutex::new(object)))
    }

    fn method(iterator: &Value, name: &str) -> Value {
        let Value::Object(object) = iterator else { panic!("iterator object") };
        object.lock().unwrap().properties.get(name).unwrap().clone()
    }

    fn check_next(vm: &mut VM, iterator: &Value, expected: i32) {
        let next = method(iterator, "next");
        let expected_index = vm.resolve_host_function_index("ecma:iterator", "next").unwrap();
        let Value::Object(function) = &next else { panic!("next function") };
        assert!(matches!(function.lock().unwrap().kind, ObjectKind::HostFunction(index) if index == expected_index));
        let result = vm.invoke(&next, std::slice::from_ref(iterator)).unwrap();
        let Value::Object(result) = result else { panic!("iterator result") };
        let result = result.lock().unwrap();
        assert!(matches!(result.properties.get("done"), Some(Value::Bool(false))));
        assert!(matches!(result.properties.get("value"), Some(Value::I32(value)) if *value == expected));
    }

    #[test]
    fn caches_use_each_vms_registry_and_reuse_wrappers_locally() {
        let mut first = VM::new();
        register(&mut first);
        let from_first = host_function(&first, "from");
        let source = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::I32(7), Value::I32(42)]))));
        let first_iterator = first.invoke(&from_first, std::slice::from_ref(&source)).unwrap();
        let first_twin = first.invoke(&from_first, std::slice::from_ref(&source)).unwrap();
        let mut second = VM::new();
        for _ in 0..3 {
            second.register_host_fn("ecma:cache-gate", "padding", Box::new(|_, _| Value::Undefined));
        }
        register(&mut second);
        assert_ne!(first.resolve_host_function_index("ecma:iterator", "next"), second.resolve_host_function_index("ecma:iterator", "next"));
        let from_second = host_function(&second, "from");
        let second_iterator = second.invoke(&from_second, std::slice::from_ref(&source)).unwrap();
        check_next(&mut first, &first_iterator, 7);
        check_next(&mut second, &second_iterator, 7);
        check_next(&mut first, &first_iterator, 42);
        let (Value::Object(first_method), Value::Object(twin_method), Value::Object(second_method)) = (
            method(&first_iterator, "next"), method(&first_twin, "next"), method(&second_iterator, "next"),
        ) else { panic!("method wrappers") };
        assert!(Arc::ptr_eq(&first_method, &twin_method));
        assert!(!Arc::ptr_eq(&first_method, &second_method));
    }
}
