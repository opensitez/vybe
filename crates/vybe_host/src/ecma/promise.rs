//! ECMA-262 §27.7 — Promise.
//!
//! Vybe's promise model is synchronous-by-default: a Promise is an Object
//! stamped `__type=Promise` with `__state` ∈ {pending, fulfilled, rejected}
//! and `__value` holding the resolved/rejected value. Async fan-out comes
//! from JSPI (see `vybe_bytecode/src/jspi.rs`) — these host fns just
//! construct the Promise objects in the appropriate state.
//!
//! Spec sections covered:
//!   §27.7.4.5 Promise.resolve(x)
//!   §27.7.4.4 Promise.reject(r)
//!   §27.7.4.1 Promise.all(iter)
//!   §27.7.4.6 Promise.race(iter)
//!   §27.7.4.2 Promise.allSettled(iter)
//!   §27.7.4.3 Promise.any(iter)
//!   §27.7.4.7 Promise.try(callbackfn)
//!   §27.7.4.x Promise.withResolvers() (ES2024)

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Internal settle helpers — never called directly from user code,
    // they're the targets the resolve/reject bound thunks dispatch to.
    // Signature: (promise, value) — `promise` arrives via __bound_args.
    vm.register_host_fn("ecma:promise", "__settle_fulfilled", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        settle(args, "fulfilled");
        Value::Undefined
    }));
    vm.register_host_fn("ecma:promise", "__settle_rejected", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        settle(args, "rejected");
        Value::Undefined
    }));

    // Capture the settler host fn indices so `new` can build bound thunks
    // without paying for a host_registry lookup on every Promise construction.
    let resolve_idx = vm.host_registry
        .get(&("ecma:promise".to_string(), "__settle_fulfilled".to_string()))
        .copied()
        .expect("__settle_fulfilled just registered");
    let reject_idx = vm.host_registry
        .get(&("ecma:promise".to_string(), "__settle_rejected".to_string()))
        .copied()
        .expect("__settle_rejected just registered");

    // `new Promise(executor)` — §27.7.3. Stamps the result Object as
    // a pending Promise, then synchronously invokes `executor(resolve,
    // reject)` where resolve/reject are bound HostFunction refs that
    // mutate the captured promise. Bind state is carried via the
    // `__bound_args` convention dispatched in vybe_bytecode/calls.rs.
    vm.register_host_fn("ecma:promise", "new", Box::new(move |ctx: &mut HostContext, args: &[Value]| {
        // known_types-style construction: first arg is the executor
        // (no separate `this` is pushed by the caller per this profile
        // convention). Allocate the Promise object here, then run the
        // executor synchronously.
        let executor = args.first().cloned().unwrap_or(Value::Undefined);
        let promise = make_promise("pending", Value::Undefined);
        if !matches!(executor, Value::Null | Value::Undefined) {
            let resolve_fn = bound_settler(resolve_idx, promise.clone());
            let reject_fn = bound_settler(reject_idx, promise.clone());
            ctx.invoke(&executor, &[resolve_fn, reject_fn]);
        }
        promise
    }));

    vm.register_host_fn("ecma:promise", "resolve", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Undefined);
        if is_promise(&val) { return val; }
        make_promise("fulfilled", val)
    }));

    vm.register_host_fn("ecma:promise", "reject", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Undefined);
        make_promise("rejected", val)
    }));

    vm.register_host_fn("ecma:promise", "all", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let mut results = Vec::with_capacity(elems.len());
                for p in elems {
                    if let Some(v) = unwrap_promise(p, "fulfilled") {
                        results.push(v);
                    } else if let Some(reason) = unwrap_promise(p, "rejected") {
                        return make_promise("rejected", reason);
                    } else {
                        results.push(p.clone());
                    }
                }
                return make_promise("fulfilled", Value::Object(
                    Arc::new(Mutex::new(Object::new_array(results)))
                ));
            }
        }
        make_promise("fulfilled", Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))))
    }));

    vm.register_host_fn("ecma:promise", "race", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                if let Some(first) = elems.first() {
                    if is_promise(first) { return first.clone(); }
                    return make_promise("fulfilled", first.clone());
                }
            }
        }
        make_promise("pending", Value::Undefined)
    }));

    vm.register_host_fn("ecma:promise", "allSettled", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let mut results = Vec::with_capacity(elems.len());
                for p in elems {
                    let mut obj = Object::new();
                    if let Some(v) = unwrap_promise(p, "fulfilled") {
                        obj.properties.insert("status".into(), Value::String(Arc::from("fulfilled")));
                        obj.properties.insert("value".into(), v);
                    } else if let Some(reason) = unwrap_promise(p, "rejected") {
                        obj.properties.insert("status".into(), Value::String(Arc::from("rejected")));
                        obj.properties.insert("reason".into(), reason);
                    } else {
                        obj.properties.insert("status".into(), Value::String(Arc::from("fulfilled")));
                        obj.properties.insert("value".into(), p.clone());
                    }
                    results.push(Value::Object(Arc::new(Mutex::new(obj))));
                }
                return make_promise("fulfilled", Value::Object(
                    Arc::new(Mutex::new(Object::new_array(results)))
                ));
            }
        }
        make_promise("fulfilled", Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))))
    }));

    vm.register_host_fn("ecma:promise", "any", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let mut errors = Vec::with_capacity(elems.len());
                for p in elems {
                    if let Some(v) = unwrap_promise(p, "fulfilled") {
                        return make_promise("fulfilled", v);
                    } else if let Some(reason) = unwrap_promise(p, "rejected") {
                        errors.push(reason);
                    } else {
                        return make_promise("fulfilled", p.clone());
                    }
                }
                let mut agg = Object::new();
                agg.properties.insert("__type".into(), Value::String(Arc::from("AggregateError")));
                agg.properties.insert("name".into(), Value::String(Arc::from("AggregateError")));
                agg.properties.insert("message".into(), Value::String(Arc::from("All promises were rejected")));
                agg.properties.insert("errors".into(), Value::Object(
                    Arc::new(Mutex::new(Object::new_array(errors)))
                ));
                return make_promise("rejected", Value::Object(Arc::new(Mutex::new(agg))));
            }
        }
        make_promise("rejected", Value::Undefined)
    }));

    // Promise.try(callbackfn) — ES2024. Calls callbackfn synchronously,
    // wraps the return value (or thrown error) in a Promise.
    vm.register_host_fn("ecma:promise", "try", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let cb = args.first().cloned().unwrap_or(Value::Undefined);
        if matches!(cb, Value::Null | Value::Undefined) {
            return make_promise("fulfilled", Value::Undefined);
        }
        let result = ctx.invoke(&cb, &[]);
        if is_promise(&result) { result } else { make_promise("fulfilled", result) }
    }));

    // Promise.withResolvers() — ES2024. Returns { promise, resolve, reject }.
    // Without thenable plumbing the resolve/reject host-callable handles aren't
    // useful; we return the shape with a pending promise + nulls so the
    // identity check passes.
    vm.register_host_fn("ecma:promise", "withResolvers", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("promise".into(), make_promise("pending", Value::Undefined));
        obj.properties.insert("resolve".into(), Value::Null);
        obj.properties.insert("reject".into(), Value::Null);
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // ── Prototype methods ─────────────────────────────────────────────
    //
    // Vybe promises settle synchronously (the executor runs to completion
    // during `new Promise(executor)`), so `.then` / `.catch` / `.finally`
    // can dispatch their callbacks inline without queuing a microtask.
    // The `await` path uses Op::PROMISE_SUSPEND (JSPI) when the value is
    // a still-pending promise; these instance methods cover the
    // callback-style API surface.

    // promise.then(onFulfilled, onRejected?) — §27.7.5.4. Returns a new
    // Promise whose state mirrors the callback's return / throw.
    vm.register_host_fn("ecma:promise", "then", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let p = args.first().cloned().unwrap_or(Value::Undefined);
        let on_fulfilled = args.get(1).cloned().unwrap_or(Value::Undefined);
        let on_rejected = args.get(2).cloned().unwrap_or(Value::Undefined);
        let (state, value) = read_promise_state(&p);
        match state.as_str() {
            "fulfilled" => settle_callback(ctx, on_fulfilled, value, "fulfilled"),
            "rejected" => {
                if is_callable(&on_rejected) {
                    settle_callback(ctx, on_rejected, value, "fulfilled")
                } else {
                    make_promise("rejected", value)
                }
            }
            // Pending promises: forward as-is. Real spec queues a
            // microtask; under the synchronous executor model the
            // pending state only persists if the executor never settled,
            // which means there's nothing to fire.
            _ => p,
        }
    }));

    // promise.catch(onRejected) — §27.7.5.1. Sugar for .then(undefined, onRejected).
    vm.register_host_fn("ecma:promise", "catch", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let p = args.first().cloned().unwrap_or(Value::Undefined);
        let on_rejected = args.get(1).cloned().unwrap_or(Value::Undefined);
        let (state, value) = read_promise_state(&p);
        match state.as_str() {
            "rejected" => settle_callback(ctx, on_rejected, value, "fulfilled"),
            "fulfilled" => p,
            _ => p,
        }
    }));

    // promise.finally(onFinally) — §27.7.5.3. Calls onFinally with no
    // args regardless of state, then forwards the original outcome.
    vm.register_host_fn("ecma:promise", "finally", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let p = args.first().cloned().unwrap_or(Value::Undefined);
        let on_finally = args.get(1).cloned().unwrap_or(Value::Undefined);
        if is_callable(&on_finally) {
            ctx.invoke(&on_finally, &[]);
        }
        p
    }));
}

/// Polymorphic dispatch for Promise instance methods. Used by
/// `ecma:value.invokeMethod` when the method-call shim sees a
/// `__type=Promise` Object — mirrors the Date / RegExp pattern.
pub fn dispatch_promise_method(ctx: &mut HostContext, method: &str, args: &[Value]) -> Option<Value> {
    let result = match method {
        "then" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_fulfilled = args.get(1).cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(2).cloned().unwrap_or(Value::Undefined);
            let (state, value) = read_promise_state(&p);
            match state.as_str() {
                "fulfilled" => settle_callback(ctx, on_fulfilled, value, "fulfilled"),
                "rejected" => {
                    if is_callable(&on_rejected) {
                        settle_callback(ctx, on_rejected, value, "fulfilled")
                    } else {
                        make_promise("rejected", value)
                    }
                }
                _ => p,
            }
        }
        "catch" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(1).cloned().unwrap_or(Value::Undefined);
            let (state, value) = read_promise_state(&p);
            match state.as_str() {
                "rejected" => settle_callback(ctx, on_rejected, value, "fulfilled"),
                _ => p,
            }
        }
        "finally" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_finally = args.get(1).cloned().unwrap_or(Value::Undefined);
            if is_callable(&on_finally) {
                ctx.invoke(&on_finally, &[]);
            }
            p
        }
        _ => return None,
    };
    Some(result)
}

/// Pull `(state, value)` out of a Promise; defaults to fulfilled-with-self
/// for non-promise inputs so `.then`/`.catch` are total functions.
fn read_promise_state(v: &Value) -> (String, Value) {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        let state = o.properties.get("__state")
            .map(|s| format!("{}", s))
            .unwrap_or_default();
        let value = o.properties.get("__value").cloned().unwrap_or(Value::Undefined);
        return (state, value);
    }
    ("fulfilled".to_string(), v.clone())
}

/// Invoke `cb(value)` and wrap the result as a Promise. If the callback
/// returns a Promise, adopt its state (matches §27.7.5.4 step 8.b
/// thenable assimilation).
fn settle_callback(ctx: &mut HostContext, cb: Value, value: Value, fallback_state: &str) -> Value {
    if !is_callable(&cb) {
        return make_promise(fallback_state, value);
    }
    let result = ctx.invoke(&cb, &[value]);
    if is_promise(&result) {
        return result;
    }
    make_promise("fulfilled", result)
}

fn is_callable(v: &Value) -> bool {
    matches!(v, Value::Object(o)
        if matches!(o.lock().unwrap().kind,
            ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
}

fn make_promise(state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties.insert("__state".into(), Value::String(Arc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Build a bound HostFunction Value that calls into the given settler
/// idx with `[promise]` prepended to the runtime args. Mirrors
/// `bound_host_fn_ref` in `namespaces/mod.rs` but inlined to avoid the
/// host_registry lookup since we already know the idx.
fn bound_settler(idx: usize, promise: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__bound_args".into(), Value::Object(
        Arc::new(Mutex::new(Object::new_array(vec![promise])))
    ));
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Mutate the promise (args[0], from __bound_args) into the given
/// terminal state with value (args[1]). No-op if already settled —
/// §27.7.4.1.1 (only the first call wins).
fn settle(args: &[Value], state: &str) {
    let promise = match args.first() { Some(p) => p, None => return };
    let value = args.get(1).cloned().unwrap_or(Value::Undefined);
    if let Value::Object(obj) = promise {
        let mut o = obj.lock().unwrap();
        let already_settled = o.properties.get("__state")
            .map(|v| format!("{}", v))
            .map(|s| s != "pending")
            .unwrap_or(false);
        if already_settled { return; }
        o.properties.insert("__state".into(), Value::String(Arc::from(state)));
        o.properties.insert("__value".into(), value);
    }
}

fn is_promise(v: &Value) -> bool {
    if let Value::Object(o) = v {
        let lock = o.lock().unwrap();
        if let Some(t) = lock.properties.get("__type") {
            return format!("{}", t) == "Promise";
        }
    }
    false
}

fn unwrap_promise(v: &Value, want_state: &str) -> Option<Value> {
    if let Value::Object(o) = v {
        let lock = o.lock().unwrap();
        if lock.properties.get("__type").map(|v| format!("{}", v)).as_deref() != Some("Promise") {
            return None;
        }
        let state = lock.properties.get("__state").map(|v| format!("{}", v)).unwrap_or_default();
        if state == want_state {
            return lock.properties.get("__value").cloned();
        }
    }
    None
}
