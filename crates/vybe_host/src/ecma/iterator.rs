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

use std::sync::{Arc, Mutex};
use crate::namespaces::receiver_host_fn_ref;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

// Method indices captured at register time so `make_iterator` can attach
// them as direct properties on every iterator instance — chained
// `Iterator.range(0,5).map(...).filter(...).toArray()` works without
// TypeRegistry vtable dispatch (the iterator object's type_id stays at
// the default Object id; method dispatch falls back to property lookup).
static METHODS: std::sync::OnceLock<Vec<(String, usize)>> = std::sync::OnceLock::new();

fn make_iterator(values: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("Iterator")));
    obj.kind = ObjectKind::Array(values);
    obj.properties.insert("__index".into(), Value::I32(0));
    if let Some(methods) = METHODS.get() {
        for (name, idx) in methods {
            obj.properties.insert(name.clone(), receiver_host_fn_ref("ecma:iterator", name, *idx));
        }
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

pub(crate) fn maybe_await_value(value: Value) -> Value {
    crate::ecma::object::unwrap_fulfilled_promise(value)
}

fn values_from_array_like(obj: &Arc<Mutex<Object>>) -> Option<Vec<Value>> {
    let o = obj.lock().unwrap();
    let ObjectKind::Array(ref vec) = o.kind else {
        return None;
    };
    let start = o.properties.get("__index").map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
    Some(vec.iter().skip(start).cloned().collect())
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

pub(crate) fn materialize_iterable_values(
    ctx: &mut HostContext,
    value: &Value,
    prefer_async: bool,
) -> Vec<Value> {
    match value {
        Value::Object(obj) => {
            if let Some(values) = values_from_array_like(obj) {
                return values;
            }
            let first = if prefer_async { "asyncIterator" } else { "iterator" };
            let second = if prefer_async { "iterator" } else { "asyncIterator" };
            if let Some(values) = crate::ecma::object::collect_protocol_iterable(ctx, obj, first) {
                return values_from_materialized(values);
            }
            if let Some(values) = crate::ecma::object::collect_protocol_iterable(ctx, obj, second) {
                return values_from_materialized(values);
            }
            Vec::new()
        }
        Value::String(text) => text
            .chars()
            .map(|ch| Value::String(Arc::from(ch.to_string().as_str())))
            .collect(),
        _ => Vec::new(),
    }
}

pub fn register(vm: &mut VM) {
    // Iterator.from(obj) — adapts arrays, iterables, generators.
    vm.register_host_fn("ecma:iterator", "from", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = args.first().cloned().unwrap_or(Value::Null);
        make_iterator(materialize_iterable_values(ctx, &v, false))
    }));

    vm.register_host_fn("ecma:iterator", "asyncFrom", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = args.first().cloned().unwrap_or(Value::Null);
        make_iterator(materialize_iterable_values(ctx, &v, true))
    }));

    // Iterator.range(start, end?, step?) — lazy numeric sequence.
    //
    // Iterator.range(n)        → 0 to n-1
    // Iterator.range(s, e)     → s to e-1
    // Iterator.range(s, e, st) → s, s+st, s+2*st, ... < e (or > e if st < 0)
    vm.register_host_fn("ecma:iterator", "range", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (start, end, step) = match args.len() {
            0 => return make_iterator(Vec::new()),
            1 => (0.0, args[0].as_f64(), 1.0),
            2 => (args[0].as_f64(), args[1].as_f64(), 1.0),
            _ => (args[0].as_f64(), args[1].as_f64(), args[2].as_f64()),
        };
        if step == 0.0 { return make_iterator(Vec::new()); }
        let mut values = Vec::new();
        let mut i = start;
        if step > 0.0 {
            while i < end { values.push(Value::F64(i)); i += step; }
        } else {
            while i > end { values.push(Value::F64(i)); i += step; }
        }
        make_iterator(values)
    }));

    // iterator.take(n) — first n elements.
    vm.register_host_fn("ecma:iterator", "take", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
        let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        make_iterator(v.into_iter().take(n).collect())
    }));

    // iterator.drop(n) — skip first n elements.
    vm.register_host_fn("ecma:iterator", "drop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
        let n = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        make_iterator(v.into_iter().skip(n).collect())
    }));

    // iterator.map(fn) — apply mapper to each.
    vm.register_host_fn("ecma:iterator", "map", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let mapper = args.get(1).cloned().unwrap_or(Value::Null);
        let mapped: Vec<Value> = v.into_iter()
            .map(|x| ctx.invoke(&mapper, &[x]))
            .collect();
        make_iterator(mapped)
    }));

    // iterator.filter(fn) — keep elements where fn returns truthy.
    vm.register_host_fn("ecma:iterator", "filter", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let pred = args.get(1).cloned().unwrap_or(Value::Null);
        let filtered: Vec<Value> = v.into_iter()
            .filter(|x| ctx.invoke(&pred, &[x.clone()]).as_bool())
            .collect();
        make_iterator(filtered)
    }));

    // iterator.reduce(fn, init?) — fold left with optional initial value.
    vm.register_host_fn("ecma:iterator", "reduce", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let reducer = args.get(1).cloned().unwrap_or(Value::Null);
        let init = args.get(2).cloned();
        let mut iter = v.into_iter();
        let mut acc = match init {
            Some(i) => i,
            None => match iter.next() { Some(x) => x, None => return Value::Undefined },
        };
        for x in iter {
            acc = ctx.invoke(&reducer, &[acc, x]);
        }
        acc
    }));

    // iterator.forEach(fn) — invoke fn for each, no result.
    vm.register_host_fn("ecma:iterator", "forEach", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let cb = args.get(1).cloned().unwrap_or(Value::Null);
        for x in v {
            ctx.invoke(&cb, &[x]);
        }
        Value::Undefined
    }));

    // iterator.some(fn) — any element matches.
    vm.register_host_fn("ecma:iterator", "some", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let pred = args.get(1).cloned().unwrap_or(Value::Null);
        Value::Bool(v.into_iter().any(|x| ctx.invoke(&pred, &[x]).as_bool()))
    }));

    // iterator.every(fn) — all elements match.
    vm.register_host_fn("ecma:iterator", "every", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let pred = args.get(1).cloned().unwrap_or(Value::Null);
        Value::Bool(v.into_iter().all(|x| ctx.invoke(&pred, &[x]).as_bool()))
    }));

    // iterator.find(fn) — first matching element, or undefined.
    vm.register_host_fn("ecma:iterator", "find", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let pred = args.get(1).cloned().unwrap_or(Value::Null);
        v.into_iter()
            .find(|x| ctx.invoke(&pred, &[x.clone()]).as_bool())
            .unwrap_or(Value::Undefined)
    }));

    // iterator.toArray() — materialize.
    vm.register_host_fn("ecma:iterator", "toArray", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(_ctx, args.first().unwrap_or(&Value::Null), false);
        Value::Object(Arc::new(Mutex::new(Object::new_array(v))))
    }));

    // iterator.flatMap(fn) — map then flatten one level.
    vm.register_host_fn("ecma:iterator", "flatMap", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let v = materialize_iterable_values(ctx, args.first().unwrap_or(&Value::Null), false);
        let mapper = args.get(1).cloned().unwrap_or(Value::Null);
        let mut result = Vec::new();
        for x in v {
            let mapped = ctx.invoke(&mapper, &[x]);
            result.extend(materialize_iterable_values(ctx, &mapped, false));
        }
        make_iterator(result)
    }));

    vm.register_host_fn("ecma:iterator", "next", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(it)) = args.first() else {
            let mut result = Object::new();
            result.properties.insert("value".into(), Value::Undefined);
            result.properties.insert("done".into(), Value::Bool(true));
            return Value::Object(Arc::new(Mutex::new(result)));
        };
        let mut lock = it.lock().unwrap();
        let index = lock.properties.get("__index").map(|value| value.as_i32().max(0) as usize).unwrap_or(0);
        if let ObjectKind::Array(ref values) = lock.kind {
            if let Some(value) = values.get(index).cloned() {
                lock.properties.insert("__index".into(), Value::I32(index as i32 + 1));
                let mut result = Object::new();
                result.properties.insert("value".into(), value);
                result.properties.insert("done".into(), Value::Bool(false));
                return Value::Object(Arc::new(Mutex::new(result)));
            }
        }
        let mut result = Object::new();
        result.properties.insert("value".into(), Value::Undefined);
        result.properties.insert("done".into(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(result)))
    }));

    // Capture method indices for instance-property attachment.
    let methods: Vec<(String, usize)> = ["next", "take", "drop", "map", "filter", "reduce",
        "forEach", "some", "every", "find", "toArray", "flatMap"]
        .iter()
        .filter_map(|name| {
            vm.host_registry
                .get(&("ecma:iterator".to_string(), name.to_string()))
                .copied()
                .map(|idx| (name.to_string(), idx))
        })
        .collect();
    let _ = METHODS.set(methods);
}
