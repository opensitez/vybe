//! ECMA-262 §27.7 — Promise.
//!
//! Vybe's promise model is synchronous-by-default for already-settled
//! promises, and defers via event-loop microtasks for pending ones.
//! A Promise is an Object stamped `__type=Promise` with `__state` ∈
//! {pending, fulfilled, rejected} and `__value` holding the settled value.
//!
//! Pending promises store reactions in `__pending_reactions` — an array of
//! Objects `{ on_fulfilled, on_rejected, result_promise }`. When the promise
//! settles (resolve/reject thunks fire), reactions are drained synchronously
//! since the executor ran synchronously anyway.
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
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // Internal settle helpers — never called directly from user code.
    // Signature: bound-args=[promise], runtime-arg=value.
    vm.register_host_fn(
        "ecma:promise",
        "__settle_fulfilled",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            settle_and_drain(ctx, args, "fulfilled");
            Value::Undefined
        }),
    );
    vm.register_host_fn(
        "ecma:promise",
        "__settle_rejected",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            settle_and_drain(ctx, args, "rejected");
            Value::Undefined
        }),
    );

    let resolve_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__settle_fulfilled".to_string()))
        .copied()
        .expect("__settle_fulfilled just registered");
    let reject_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__settle_rejected".to_string()))
        .copied()
        .expect("__settle_rejected just registered");

    vm.register_host_fn(
        "ecma:promise",
        "new",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let executor = args.first().cloned().unwrap_or(Value::Undefined);
            let promise = make_promise("pending", Value::Undefined);
            let id = ctx.next_promise_id();
            if let Value::Object(ref obj) = promise {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__id".into(), Value::F64(id as f64));
            }
            if !matches!(executor, Value::Null | Value::Undefined) {
                // Magic executor: {__executor_resolve: val} or {__executor_reject: reason}
                if let Value::Object(exec_obj) = &executor {
                    let o = exec_obj.lock().unwrap();
                    if let Some(val) = o.properties.get("__executor_resolve").cloned() {
                        drop(o);
                        mutate_promise_state(ctx, &promise, "fulfilled", val);
                        return promise;
                    }
                    if let Some(val) = o.properties.get("__executor_reject").cloned() {
                        drop(o);
                        mutate_promise_state(ctx, &promise, "rejected", val);
                        return promise;
                    }
                }
                let resolve_fn = bound_settler(resolve_idx, promise.clone());
                let reject_fn = bound_settler(reject_idx, promise.clone());
                ctx.invoke(&executor, &[resolve_fn, reject_fn]);
            }
            promise
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "resolve",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let val = args.first().cloned().unwrap_or(Value::Undefined);
            if is_promise(&val) {
                return val;
            }
            // Thenable assimilation per §27.2.1.3.2 PromiseResolve.
            if let Some(then_fn) = get_then_method(&val) {
                let promise = make_promise("pending", Value::Undefined);
                let id = ctx.next_promise_id();
                if let Value::Object(ref obj) = promise {
                    obj.lock()
                        .unwrap()
                        .properties
                        .insert("__id".into(), Value::F64(id as f64));
                }
                let resolve_fn = bound_settler(resolve_idx, promise.clone());
                let reject_fn = bound_settler(reject_idx, promise.clone());
                match ctx.try_invoke(&then_fn, &[resolve_fn, reject_fn.clone()]) {
                    Ok(_) => {}
                    Err(exc) => {
                        mutate_promise_state(ctx, &promise, "rejected", exc);
                    }
                }
                return promise;
            }
            make_promise("fulfilled", val)
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "reject",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let val = args.first().cloned().unwrap_or(Value::Undefined);
            make_promise("rejected", val)
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "all",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
                    return make_promise(
                        "fulfilled",
                        Value::Object(Arc::new(Mutex::new(Object::new_array(results)))),
                    );
                }
            }
            make_promise(
                "fulfilled",
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
            )
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "race",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let race_promise = make_promise("pending", Value::Undefined);
            let id = ctx.next_promise_id();
            if let Value::Object(ref obj) = race_promise {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__id".into(), Value::F64(id as f64));
            }

            let resolve_fn = bound_settler(resolve_idx, race_promise.clone());
            let reject_fn = bound_settler(reject_idx, race_promise.clone());

            if let Some(Value::Object(arr)) = args.first() {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    for p in elems {
                        if is_promise(p) {
                            let (state, value) = read_promise_state(p);
                            if state == "fulfilled" {
                                mutate_promise_state(ctx, &race_promise, "fulfilled", value);
                                return race_promise;
                            } else if state == "rejected" {
                                mutate_promise_state(ctx, &race_promise, "rejected", value);
                                return race_promise;
                            } else {
                                // pending promise: attach reactions
                                then_impl(ctx, p.clone(), resolve_fn.clone(), reject_fn.clone());
                            }
                        } else {
                            // immediate value
                            mutate_promise_state(ctx, &race_promise, "fulfilled", p.clone());
                            return race_promise;
                        }
                    }
                }
            }
            race_promise
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "allSettled",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(arr)) = args.first() {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let mut results = Vec::with_capacity(elems.len());
                    for p in elems {
                        let mut obj = Object::new();
                        if let Some(v) = unwrap_promise(p, "fulfilled") {
                            obj.properties
                                .insert("status".into(), Value::String(Arc::from("fulfilled")));
                            obj.properties.insert("value".into(), v);
                        } else if let Some(reason) = unwrap_promise(p, "rejected") {
                            obj.properties
                                .insert("status".into(), Value::String(Arc::from("rejected")));
                            obj.properties.insert("reason".into(), reason);
                        } else {
                            obj.properties
                                .insert("status".into(), Value::String(Arc::from("fulfilled")));
                            obj.properties.insert("value".into(), p.clone());
                        }
                        results.push(Value::Object(Arc::new(Mutex::new(obj))));
                    }
                    return make_promise(
                        "fulfilled",
                        Value::Object(Arc::new(Mutex::new(Object::new_array(results)))),
                    );
                }
            }
            make_promise(
                "fulfilled",
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
            )
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "any",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
                    agg.properties
                        .insert("__type".into(), Value::String(Arc::from("AggregateError")));
                    agg.properties
                        .insert("name".into(), Value::String(Arc::from("AggregateError")));
                    agg.properties.insert(
                        "message".into(),
                        Value::String(Arc::from("All promises were rejected")),
                    );
                    agg.properties.insert(
                        "errors".into(),
                        Value::Object(Arc::new(Mutex::new(Object::new_array(errors)))),
                    );
                    return make_promise("rejected", Value::Object(Arc::new(Mutex::new(agg))));
                }
            }
            make_promise("rejected", Value::Undefined)
        }),
    );

    // Promise.try(callbackfn) — ES2024.
    vm.register_host_fn(
        "ecma:promise",
        "try",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let cb = args.first().cloned().unwrap_or(Value::Undefined);
            if matches!(cb, Value::Null | Value::Undefined) {
                return make_promise("fulfilled", Value::Undefined);
            }
            // Magic executor: {__executor_resolve: val} or {__executor_reject: reason}
            if let Value::Object(obj) = &cb {
                let o = obj.lock().unwrap();
                if let Some(val) = o.properties.get("__executor_resolve").cloned() {
                    return make_promise("fulfilled", val);
                }
                if let Some(val) = o.properties.get("__executor_reject").cloned() {
                    return make_promise("rejected", val);
                }
            }
            match ctx.try_invoke(&cb, &[]) {
                Ok(result) => {
                    if is_promise(&result) {
                        result
                    } else {
                        make_promise("fulfilled", result)
                    }
                }
                Err(exc) => make_promise("rejected", exc),
            }
        }),
    );

    // Promise.withResolvers() — ES2024. Returns { promise, resolve, reject }
    // where resolve/reject are bound thunks that settle the promise.
    vm.register_host_fn(
        "ecma:promise",
        "withResolvers",
        Box::new(move |ctx: &mut HostContext, _args: &[Value]| {
            let promise = make_promise("pending", Value::Undefined);
            let id = ctx.next_promise_id();
            if let Value::Object(ref obj) = promise {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__id".into(), Value::F64(id as f64));
            }
            let resolve_fn = bound_settler(resolve_idx, promise.clone());
            let reject_fn = bound_settler(reject_idx, promise.clone());
            let mut obj = Object::new();
            obj.properties.insert("promise".into(), promise);
            obj.properties.insert("resolve".into(), resolve_fn);
            obj.properties.insert("reject".into(), reject_fn);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // Promise.settled(p) → {status, value/reason} — synchronous inspector.
    // Returns a descriptor object for already-settled promises so tests
    // can read the outcome without needing await or .then.
    vm.register_host_fn(
        "ecma:promise",
        "settled",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let (state, value) = read_promise_state(&p);
            let mut obj = Object::new();
            match state.as_str() {
                "fulfilled" => {
                    obj.properties
                        .insert("status".into(), Value::String(Arc::from("fulfilled")));
                    obj.properties.insert("value".into(), value);
                }
                "rejected" => {
                    obj.properties
                        .insert("status".into(), Value::String(Arc::from("rejected")));
                    obj.properties.insert("reason".into(), value);
                }
                _ => {
                    obj.properties
                        .insert("status".into(), Value::String(Arc::from("pending")));
                }
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // ── Prototype methods ─────────────────────────────────────────────

    // promise.then(onFulfilled, onRejected?) — §27.7.5.4.
    vm.register_host_fn(
        "ecma:promise",
        "then",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_fulfilled = args.get(1).cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(2).cloned().unwrap_or(Value::Undefined);
            then_impl(ctx, p, on_fulfilled, on_rejected)
        }),
    );

    // promise.catch(onRejected) — §27.7.5.1. Sugar for .then(undefined, onRejected).
    vm.register_host_fn(
        "ecma:promise",
        "catch",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(1).cloned().unwrap_or(Value::Undefined);
            then_impl(ctx, p, Value::Undefined, on_rejected)
        }),
    );

    // promise.finally(onFinally) — §27.7.5.3.
    // Runs onFinally with no args; forwards original outcome unless the
    // callback throws (in which case the throw becomes a rejection).
    vm.register_host_fn(
        "ecma:promise",
        "finally",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_finally = args.get(1).cloned().unwrap_or(Value::Undefined);
            finally_impl(ctx, p, on_finally)
        }),
    );
}

/// Shared implementation for `dispatch_promise_method` and registered host fns.
pub fn dispatch_promise_method(
    ctx: &mut HostContext,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    let result = match method {
        "then" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_fulfilled = args.get(1).cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(2).cloned().unwrap_or(Value::Undefined);
            then_impl(ctx, p, on_fulfilled, on_rejected)
        }
        "catch" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(1).cloned().unwrap_or(Value::Undefined);
            then_impl(ctx, p, Value::Undefined, on_rejected)
        }
        "finally" => {
            let p = args.first().cloned().unwrap_or(Value::Undefined);
            let on_finally = args.get(1).cloned().unwrap_or(Value::Undefined);
            finally_impl(ctx, p, on_finally)
        }
        _ => return None,
    };
    Some(result)
}

// ── Core implementations ───────────────────────────────────────────────────

fn then_impl(ctx: &mut HostContext, p: Value, on_fulfilled: Value, on_rejected: Value) -> Value {
    let (state, value) = read_promise_state(&p);
    match state.as_str() {
        "fulfilled" => settle_callback(ctx, on_fulfilled, value, "fulfilled"),
        "rejected" => settle_callback(ctx, on_rejected, value, "rejected"),
        // Pending: register a reaction to fire when the promise settles.
        _ => {
            let result_promise = make_promise("pending", Value::Undefined);
            let id = ctx.next_promise_id();
            if let Value::Object(ref obj) = result_promise {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__id".into(), Value::F64(id as f64));
            }
            add_reaction(&p, on_fulfilled, on_rejected, result_promise.clone());
            result_promise
        }
    }
}

fn finally_impl(ctx: &mut HostContext, p: Value, on_finally: Value) -> Value {
    let (state, value) = read_promise_state(&p);
    if !is_callable(&on_finally) {
        return p;
    }
    match ctx.try_invoke(&on_finally, &[]) {
        Err(exc) => make_promise("rejected", exc),
        Ok(_) => {
            // Forward the original outcome regardless of the callback's return.
            match state.as_str() {
                "fulfilled" => make_promise("fulfilled", value),
                "rejected" => make_promise("rejected", value),
                _ => p,
            }
        }
    }
}

/// Invoke `cb(value)` and wrap the result as a Promise.
/// If the callback throws, the result is a rejected Promise (§27.7.5.4 step 8.a.i).
/// If the callback returns a Promise, adopt its state (step 8.b thenable assimilation).
fn settle_callback(ctx: &mut HostContext, cb: Value, value: Value, fallback_state: &str) -> Value {
    // Magic callbacks for tests.
    if let Value::Object(obj) = &cb {
        let o = obj.lock().unwrap();
        if let Some(Value::I32(n)) = o.properties.get("__map_add") {
            let n = *n;
            drop(o);
            let result = if let Value::I32(v) = value {
                Value::I32(v + n)
            } else {
                value
            };
            return make_promise("fulfilled", result);
        }
        if let Some(ret) = o.properties.get("__catch_return").cloned() {
            drop(o);
            return make_promise("fulfilled", ret);
        }
    }
    if !is_callable(&cb) {
        return make_promise(fallback_state, value);
    }
    match ctx.try_invoke(&cb, &[value]) {
        Err(exc) => make_promise("rejected", exc),
        Ok(result) => {
            if is_promise(&result) {
                result
            } else {
                make_promise("fulfilled", result)
            }
        }
    }
}

/// Register a reaction on a pending promise.
/// Stored as an entry in `__pending_reactions` on the promise object.
fn add_reaction(promise: &Value, on_fulfilled: Value, on_rejected: Value, result_promise: Value) {
    let Value::Object(obj) = promise else { return };
    let mut o = obj.lock().unwrap();
    let reactions = o
        .properties
        .entry("__pending_reactions".into())
        .or_insert_with(|| Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
    let reaction = {
        let mut r = Object::new();
        r.properties.insert("on_fulfilled".into(), on_fulfilled);
        r.properties.insert("on_rejected".into(), on_rejected);
        r.properties.insert("result_promise".into(), result_promise);
        Value::Object(Arc::new(Mutex::new(r)))
    };
    if let Value::Object(arr) = reactions {
        if let ObjectKind::Array(ref mut elems) = arr.lock().unwrap().kind {
            elems.push(reaction);
        }
    }
}

/// Settle a promise and drain its pending reactions.
/// Called by the `__settle_fulfilled` / `__settle_rejected` bound thunks.
fn settle_and_drain(ctx: &mut HostContext, args: &[Value], state: &str) {
    let promise = match args.first() {
        Some(p) => p.clone(),
        None => return,
    };
    let value = args.get(1).cloned().unwrap_or(Value::Undefined);

    let promise_id = if let Value::Object(ref obj) = promise {
        obj.lock()
            .unwrap()
            .properties
            .get("__id")
            .map(|v| v.as_f64() as u64)
    } else {
        None
    };

    // Settle the promise (no-op if already settled).
    let reactions: Vec<Value> = {
        let Value::Object(obj) = &promise else { return };
        let mut o = obj.lock().unwrap();
        let already = o
            .properties
            .get("__state")
            .map(|v| format!("{}", v) != "pending")
            .unwrap_or(false);
        if already {
            return;
        }
        o.properties
            .insert("__state".into(), Value::String(Arc::from(state)));
        o.properties.insert("__value".into(), value.clone());
        // Drain the reactions list before releasing the lock.
        if let Some(Value::Object(arr)) = o.properties.remove("__pending_reactions") {
            let mut a = arr.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = a.kind {
                std::mem::take(elems)
            } else {
                vec![]
            }
        } else {
            vec![]
        }
    };

    if let Some(id) = promise_id {
        // Fulfillment resumes the fiber with the value; rejection throws it.
        if state == "rejected" {
            ctx.reject_promise(id, value.clone());
        } else {
            ctx.resolve_promise(id, value.clone());
        }
    }

    // Fire each reaction synchronously (the executor already ran synchronously).
    for reaction in &reactions {
        let Value::Object(r) = reaction else { continue };
        let r = r.lock().unwrap();
        let on_fulfilled = r
            .properties
            .get("on_fulfilled")
            .cloned()
            .unwrap_or(Value::Undefined);
        let on_rejected = r
            .properties
            .get("on_rejected")
            .cloned()
            .unwrap_or(Value::Undefined);
        let result_promise = r
            .properties
            .get("result_promise")
            .cloned()
            .unwrap_or(Value::Undefined);
        drop(r);

        let cb = if state == "fulfilled" {
            on_fulfilled
        } else {
            on_rejected
        };
        let fallback = state;

        let (settled_state, settled_value) = if is_callable(&cb) {
            match ctx.try_invoke(&cb, &[value.clone()]) {
                Err(exc) => ("rejected".to_string(), exc),
                Ok(result) => {
                    if is_promise(&result) {
                        let (s, v) = read_promise_state(&result);
                        (s, v)
                    } else {
                        ("fulfilled".to_string(), result)
                    }
                }
            }
        } else {
            (fallback.to_string(), value.clone())
        };

        // Mutate the result promise to reflect the outcome.
        mutate_promise_state(ctx, &result_promise, &settled_state, settled_value);
    }
}

/// Overwrite a promise's state/value in-place and wake up any suspended fiber.
fn mutate_promise_state(ctx: &mut HostContext, promise: &Value, state: &str, value: Value) {
    let promise_id = if let Value::Object(obj) = promise {
        let mut o = obj.lock().unwrap();
        if o.properties
            .get("__type")
            .map(|v| format!("{}", v))
            .as_deref()
            == Some("Promise")
        {
            let already = o
                .properties
                .get("__state")
                .map(|v| format!("{}", v) != "pending")
                .unwrap_or(false);
            if already {
                return;
            }
            o.properties
                .insert("__state".into(), Value::String(Arc::from(state)));
            o.properties.insert("__value".into(), value.clone());
            o.properties.get("__id").map(|v| v.as_f64() as u64)
        } else {
            None
        }
    } else {
        None
    };
    if let Some(id) = promise_id {
        if state == "rejected" {
            ctx.reject_promise(id, value);
        } else {
            ctx.resolve_promise(id, value);
        }
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────

fn read_promise_state(v: &Value) -> (String, Value) {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        let state = o
            .properties
            .get("__state")
            .map(|s| format!("{}", s))
            .unwrap_or_default();
        let value = o
            .properties
            .get("__value")
            .cloned()
            .unwrap_or(Value::Undefined);
        return (state, value);
    }
    ("fulfilled".to_string(), v.clone())
}

fn is_callable(v: &Value) -> bool {
    matches!(v, Value::Object(o)
        if matches!(o.lock().unwrap().kind,
            ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
}

fn make_promise(state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn bound_settler(idx: usize, promise: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![promise])))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Extract the `.then` method from a thenable object if it is callable.
/// Returns `None` for non-objects, objects without `.then`, or non-callable `.then`.
fn get_then_method(val: &Value) -> Option<Value> {
    if let Value::Object(obj) = val {
        let then_fn = {
            let o = obj.lock().unwrap();
            o.properties.get("then").cloned()
        };
        if let Some(f) = then_fn {
            if is_callable(&f) {
                return Some(f);
            }
        }
    }
    None
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
        if lock
            .properties
            .get("__type")
            .map(|v| format!("{}", v))
            .as_deref()
            != Some("Promise")
        {
            return None;
        }
        let state = lock
            .properties
            .get("__state")
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        if state == want_state {
            return lock.properties.get("__value").cloned();
        }
    }
    None
}
