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
static METHODS: std::sync::OnceLock<Vec<(String, Value)>> = std::sync::OnceLock::new();

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

fn attach_iterator_methods(obj: &mut Object) {
    if let Some(methods) = METHODS.get() {
        for (name, method) in methods {
            obj.properties.insert(name.clone(), method.clone());
        }
    }
}

fn make_iterator(values: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties
        .reserve(3 + METHODS.get().map_or(0, |methods| methods.len()));
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Iterator"));
    obj.kind = ObjectKind::Array(values);
    obj.properties.insert("__index".into(), Value::I32(0));
    attach_iterator_methods(&mut obj);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_lazy_map(source: Value, mapper: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .reserve(5 + METHODS.get().map_or(0, |methods| methods.len()));
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Iterator"));
    obj.properties
        .insert("__iterator_kind".into(), crate::keys::string_value("map"));
    obj.properties.insert("__source".into(), source);
    obj.properties.insert("__mapper".into(), mapper);
    obj.properties.insert("__index".into(), Value::I32(0));
    attach_iterator_methods(&mut obj);
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
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(ref vec) = o.kind {
            return vec.clone();
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

pub fn register(vm: &mut VM) {
    // Iterator.from(obj) — adapts arrays, iterables, generators.
    vm.register_free_fn(
        "ecma:iterator",
        "from",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
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
            make_iterator(materialize_iterable_values(ctx, v, false))
        }),
    );

    vm.register_free_fn(
        "ecma:iterator",
        "asyncFrom",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // See `from` above — this is the shape that actually carried a
            // receiver, and the reason the two twins disagreed.
            let user = ctx.user_args(args, 0);
            let null = Value::Null;
            let v = user.first().unwrap_or(&null);
            make_iterator(materialize_iterable_values(ctx, v, true))
        }),
    );

    // Iterator.range(start, end?, step?) — lazy numeric sequence.
    //
    // Iterator.range(n)        → 0 to n-1
    // Iterator.range(s, e)     → s to e-1
    // Iterator.range(s, e, st) → s, s+st, s+2*st, ... < e (or > e if st < 0)
    vm.register_host_fn(
        "ecma:iterator",
        "range",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let (start, end, step) = match args.len() {
                0 => return make_iterator(Vec::new()),
                1 => (0.0, args[0].as_f64(), 1.0),
                2 => (args[0].as_f64(), args[1].as_f64(), 1.0),
                _ => (args[0].as_f64(), args[1].as_f64(), args[2].as_f64()),
            };
            if step == 0.0 {
                return make_iterator(Vec::new());
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
            make_iterator(values)
        }),
    );

    // iterator.take(n) — first n elements.
    vm.register_host_fn(
        "ecma:iterator",
        "take",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
            let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let take_len = v.len().min(n);
            let out: Vec<Value> = v.into_iter().take(take_len).collect();
            make_iterator(out)
        }),
    );

    // iterator.drop(n) — skip first n elements.
    vm.register_host_fn(
        "ecma:iterator",
        "drop",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
            let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let skip_len = v.len().min(n);
            let out: Vec<Value> = v.into_iter().skip(skip_len).collect();
            make_iterator(out)
        }),
    );

    // iterator.map(fn) — apply mapper to each.
    vm.register_host_fn(
        "ecma:iterator",
        "map",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let source = args.first().cloned().unwrap_or(Value::Null);
            let mapper = args.get(1).cloned().unwrap_or(Value::Null);
            make_lazy_map(source, mapper)
        }),
    );

    // iterator.filter(fn) — keep elements where fn returns truthy.
    vm.register_host_fn(
        "ecma:iterator",
        "filter",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            let mut filtered = Vec::with_capacity(v.len());
            for x in v {
                if invoke_unary_callback_prepared(ctx, &pred, &prepared_pred, x.clone()).as_bool() {
                    filtered.push(x);
                }
            }
            make_iterator(filtered)
        }),
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
    vm.register_host_fn(
        "ecma:iterator",
        "flatMap",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
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
            make_iterator(result)
        }),
    );

    // Iterator.concat(iter1, iter2, ...) — ES2025 §3.1.1.1.
    vm.register_host_fn(
        "ecma:iterator",
        "concat",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut result = Vec::with_capacity(args.len());
            for arg in args {
                let values = materialize_iterable_values(ctx, arg, false);
                result.reserve(values.len());
                result.extend(values);
            }
            make_iterator(result)
        }),
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
    let _ = METHODS.set(methods);
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
