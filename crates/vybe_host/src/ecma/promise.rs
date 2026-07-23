//! ECMA-262 §27.7 — Promise.
//!
//! Promise reactions are scheduled as event-loop microtasks, including
//! reactions attached to already-settled promises.
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

use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

static PROMISE_REACTION_HOST_IDX: OnceLock<usize> = OnceLock::new();
// Host-fn indices for the settle thunks, so free functions (resolve_promise_with_value)
// can build bound settlers to adopt a returned promise/thenable's eventual state.
static SETTLE_FULFILLED_IDX: OnceLock<usize> = OnceLock::new();
static SETTLE_REJECTED_IDX: OnceLock<usize> = OnceLock::new();
// §27.2.1.3.2 Promise Resolve Function — resolves THROUGH thenables
// (unlike __settle_fulfilled, which settles with the raw value). This is
// what the executor's `resolve` and thenable-job callbacks must be.
static RESOLVE_IDX: OnceLock<usize> = OnceLock::new();
// §27.2.2.2 NewPromiseResolveThenableJob — calls thenable.then(res, rej)
// in a MICROTASK with `this` = thenable.
static THENABLE_JOB_IDX: OnceLock<usize> = OnceLock::new();
// Settles a promise with a *forced* (state, value), ignoring the awaited value.
// Used by `.finally` to preserve the original settlement after awaiting a
// thenable the finally callback returned.
static PRESERVE_IDX: OnceLock<usize> = OnceLock::new();

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

    // Promise.all element resolver. bound-args=[aggregate, index],
    // runtime-arg=value. Stores the value at its slot and settles the
    // aggregate once every element has fulfilled.
    vm.register_host_fn(
        "ecma:promise",
        "__all_element",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let value = args.get(2).cloned().unwrap_or(Value::Undefined);
            let complete = aggregate_record_element(&aggregate, index, value);
            if let Some(results) = complete {
                settle_and_drain(ctx, &[aggregate, results], "fulfilled");
            }
            Value::Undefined
        }),
    );
    // Promise.allSettled element handlers. bound-args=[aggregate, index],
    // runtime-arg=value/reason. Store the descriptor and fulfill the aggregate
    // after every input has settled.
    vm.register_host_fn(
        "ecma:promise",
        "__allsettled_fulfilled",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let value = args.get(2).cloned().unwrap_or(Value::Undefined);
            let complete =
                aggregate_record_element(&aggregate, index, settled_descriptor("fulfilled", value));
            if let Some(results) = complete {
                settle_and_drain(ctx, &[aggregate, results], "fulfilled");
            }
            Value::Undefined
        }),
    );
    vm.register_host_fn(
        "ecma:promise",
        "__allsettled_rejected",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let reason = args.get(2).cloned().unwrap_or(Value::Undefined);
            let complete =
                aggregate_record_element(&aggregate, index, settled_descriptor("rejected", reason));
            if let Some(results) = complete {
                settle_and_drain(ctx, &[aggregate, results], "fulfilled");
            }
            Value::Undefined
        }),
    );
    // Aggregate rejecter (Promise.all / Promise.race short-circuit).
    // bound-args=[aggregate], runtime-arg=reason.
    vm.register_host_fn(
        "ecma:promise",
        "__aggregate_reject",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let reason = args.get(1).cloned().unwrap_or(Value::Undefined);
            settle_and_drain(ctx, &[aggregate, reason], "rejected");
            Value::Undefined
        }),
    );
    // Promise.any fulfillment/rejection handlers. Fulfillment short-circuits;
    // rejection records the reason and rejects with AggregateError once every
    // input has rejected.
    vm.register_host_fn(
        "ecma:promise",
        "__any_fulfilled",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let value = args.get(1).cloned().unwrap_or(Value::Undefined);
            settle_and_drain(ctx, &[aggregate, value], "fulfilled");
            Value::Undefined
        }),
    );
    vm.register_host_fn(
        "ecma:promise",
        "__any_rejected",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let aggregate = args.first().cloned().unwrap_or(Value::Undefined);
            let index = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let reason = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(error) = any_record_rejection(&aggregate, index, reason) {
                settle_and_drain(ctx, &[aggregate, error], "rejected");
            }
            Value::Undefined
        }),
    );
    vm.register_host_fn(
        "ecma:promise",
        "__reaction",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let result_promise = args.first().cloned().unwrap_or(Value::Undefined);
            let state = args
                .get(1)
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| "fulfilled".to_string());
            let on_fulfilled = args.get(2).cloned().unwrap_or(Value::Undefined);
            let on_rejected = args.get(3).cloned().unwrap_or(Value::Undefined);
            let value = args.get(4).cloned().unwrap_or(Value::Undefined);
            run_reaction(
                ctx,
                result_promise,
                &state,
                on_fulfilled,
                on_rejected,
                value,
            );
            Value::Undefined
        }),
    );

    // Adopt: resolve `target` with `source` per §27.2.1.3.2 — if `source` is a
    // (possibly pending) promise/thenable, target adopts its eventual state.
    // Used by the VM's JSPI promising boundary when an async body returns a
    // still-pending promise. args = [target, source].
    vm.register_host_fn(
        "ecma:promise",
        "__adopt",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let source = args.get(1).cloned().unwrap_or(Value::Undefined);
            resolve_promise_with_value(ctx, &target, source);
            Value::Undefined
        }),
    );

    // §27.2.1.3.2 Promise Resolve Function. bound-args=[promise],
    // runtime-arg=resolution. Resolves THROUGH promises/thenables (adopt
    // eventual state / queue a thenable job) instead of settling raw.
    vm.register_host_fn(
        "ecma:promise",
        "__resolve",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let promise = args.first().cloned().unwrap_or(Value::Undefined);
            let value = args.get(1).cloned().unwrap_or(Value::Undefined);
            resolve_promise_with_value(ctx, &promise, value);
            Value::Undefined
        }),
    );

    // §27.2.2.2 NewPromiseResolveThenableJob. bound-args=[promise,
    // thenable, then_fn]. Calls then_fn with this=thenable and the
    // promise's (resolve, reject) functions; a throw rejects.
    vm.register_host_fn(
        "ecma:promise",
        "__thenable_job",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let promise = args.first().cloned().unwrap_or(Value::Undefined);
            let thenable = args.get(1).cloned().unwrap_or(Value::Undefined);
            let then_fn = args.get(2).cloned().unwrap_or(Value::Undefined);
            let res = RESOLVE_IDX
                .get()
                .map(|&i| bound_settler(i, promise.clone()))
                .unwrap_or(Value::Undefined);
            let rej = SETTLE_REJECTED_IDX
                .get()
                .map(|&i| bound_settler(i, promise.clone()))
                .unwrap_or(Value::Undefined);
            let saved_this = ctx.current_js_this();
            ctx.set_js_this(thenable.clone());
            let outcome = ctx.try_invoke(&then_fn, &[res, rej]);
            ctx.set_js_this(saved_this);
            if let Err(exc) = outcome {
                mutate_promise_state(ctx, &promise, "rejected", exc);
            }
            Value::Undefined
        }),
    );

    // Force-settle a promise with a bound (state, value), ignoring the runtime
    // (awaited) value. bound-args = [promise, state, forced_value].
    vm.register_host_fn(
        "ecma:promise",
        "__preserve",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let promise = args.first().cloned().unwrap_or(Value::Undefined);
            let state = args
                .get(1)
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| "fulfilled".to_string());
            let forced = args.get(2).cloned().unwrap_or(Value::Undefined);
            mutate_promise_state(ctx, &promise, &state, forced);
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
    let all_element_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__all_element".to_string()))
        .copied()
        .expect("__all_element just registered");
    let allsettled_fulfilled_idx = vm
        .host_registry
        .get(&(
            "ecma:promise".to_string(),
            "__allsettled_fulfilled".to_string(),
        ))
        .copied()
        .expect("__allsettled_fulfilled just registered");
    let allsettled_rejected_idx = vm
        .host_registry
        .get(&(
            "ecma:promise".to_string(),
            "__allsettled_rejected".to_string(),
        ))
        .copied()
        .expect("__allsettled_rejected just registered");
    let aggregate_reject_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__aggregate_reject".to_string()))
        .copied()
        .expect("__aggregate_reject just registered");
    let any_fulfilled_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__any_fulfilled".to_string()))
        .copied()
        .expect("__any_fulfilled just registered");
    let any_rejected_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__any_rejected".to_string()))
        .copied()
        .expect("__any_rejected just registered");
    let reaction_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__reaction".to_string()))
        .copied()
        .expect("__reaction just registered");
    let _ = PROMISE_REACTION_HOST_IDX.set(reaction_idx);
    let _ = SETTLE_FULFILLED_IDX.set(resolve_idx);
    let _ = SETTLE_REJECTED_IDX.set(reject_idx);
    let resolve_through_idx = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__resolve".to_string()))
        .copied()
        .expect("__resolve just registered");
    let _ = RESOLVE_IDX.set(resolve_through_idx);
    if let Some(&tj) = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__thenable_job".to_string()))
    {
        let _ = THENABLE_JOB_IDX.set(tj);
    }
    if let Some(&preserve_idx) = vm
        .host_registry
        .get(&("ecma:promise".to_string(), "__preserve".to_string()))
    {
        let _ = PRESERVE_IDX.set(preserve_idx);
    }

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
                // §27.2.3.1: resolve is a Promise RESOLVE Function
                // (resolves through thenables); reject settles raw.
                let resolve_fn = bound_settler(resolve_through_idx, promise.clone());
                let reject_fn = bound_settler(reject_idx, promise.clone());
                if let Err(exc) = ctx.try_invoke(&executor, &[resolve_fn, reject_fn]) {
                    mutate_promise_state(ctx, &promise, "rejected", exc);
                }
            }
            promise
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "resolve",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let val = args.first().cloned().unwrap_or(Value::Undefined);
            // §27.2.4.7 PromiseResolve: a native promise passes through;
            // everything else resolves a fresh promise with the value —
            // resolve_promise_with_value queues the ThenableJob for
            // thenables and fulfills plain values directly.
            if is_promise(&val) {
                return val;
            }
            let promise = pending_promise_with_id(ctx);
            resolve_promise_with_value(ctx, &promise, val);
            promise
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
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let Some(inputs) = promise_combinator_inputs(ctx, args, "Promise.all") else {
                return Value::Undefined;
            };
            let n = inputs.len();
            // Aggregate promise: pending until every element fulfils (then
            // fulfils with an in-order results array) or any rejects (then
            // rejects with that reason). Pending inputs are awaited via
            // reactions so `await Promise.all([...])` resumes once all settle.
            let aggregate = make_promise("pending", Value::Undefined);
            let id = ctx.next_promise_id();
            let results = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                    Value::Undefined;
                    n
                ]))));
            if let Value::Object(ref obj) = aggregate {
                let mut o = obj.lock().unwrap();
                o.properties.insert("__id".into(), Value::F64(id as f64));
                o.properties.insert("__all_results".into(), results.clone());
                o.properties
                    .insert("__all_remaining".into(), Value::F64(n as f64));
            }
            if n == 0 {
                settle_and_drain(ctx, &[aggregate.clone(), results], "fulfilled");
                return aggregate;
            }
            for (i, input) in inputs.into_iter().enumerate() {
                let p = promise_resolve_for_combinator(ctx, input);
                let (state, value) = read_promise_state(&p);
                if state == "fulfilled" {
                    if let Some(done) = aggregate_record_element(&aggregate, i, value) {
                        settle_and_drain(ctx, &[aggregate.clone(), done], "fulfilled");
                    }
                } else if state == "rejected" {
                    settle_and_drain(ctx, &[aggregate.clone(), value], "rejected");
                    return aggregate;
                } else {
                    let on_f =
                        bound_settler2(all_element_idx, aggregate.clone(), Value::F64(i as f64));
                    let on_r = bound_settler(aggregate_reject_idx, aggregate.clone());
                    add_reaction(&p, on_f, on_r, make_promise("pending", Value::Undefined));
                }
            }
            aggregate
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "race",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let Some(inputs) = promise_combinator_inputs(ctx, args, "Promise.race") else {
                return Value::Undefined;
            };
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

            for input in inputs {
                let p = promise_resolve_for_combinator(ctx, input);
                let (state, value) = read_promise_state(&p);
                if state == "fulfilled" {
                    mutate_promise_state(ctx, &race_promise, "fulfilled", value);
                    return race_promise;
                } else if state == "rejected" {
                    mutate_promise_state(ctx, &race_promise, "rejected", value);
                    return race_promise;
                } else {
                    then_impl(ctx, p, resolve_fn.clone(), reject_fn.clone());
                }
            }
            race_promise
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "allSettled",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let Some(inputs) = promise_combinator_inputs(ctx, args, "Promise.allSettled") else {
                return Value::Undefined;
            };
            let n = inputs.len();
            let aggregate = pending_promise_with_id(ctx);
            let results = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                Value::Undefined;
                n
            ]))));
            if let Value::Object(ref obj) = aggregate {
                let mut o = obj.lock().unwrap();
                o.properties.insert("__all_results".into(), results.clone());
                o.properties
                    .insert("__all_remaining".into(), Value::F64(n as f64));
            }
            if n == 0 {
                settle_and_drain(ctx, &[aggregate.clone(), results], "fulfilled");
                return aggregate;
            }
            for (i, input) in inputs.into_iter().enumerate() {
                let p = promise_resolve_for_combinator(ctx, input);
                let (state, value) = read_promise_state(&p);
                if state == "fulfilled" {
                    if let Some(done) = aggregate_record_element(
                        &aggregate,
                        i,
                        settled_descriptor("fulfilled", value),
                    ) {
                        settle_and_drain(ctx, &[aggregate.clone(), done], "fulfilled");
                    }
                } else if state == "rejected" {
                    if let Some(done) = aggregate_record_element(
                        &aggregate,
                        i,
                        settled_descriptor("rejected", value),
                    ) {
                        settle_and_drain(ctx, &[aggregate.clone(), done], "fulfilled");
                    }
                } else {
                    let on_f = bound_settler2(
                        allsettled_fulfilled_idx,
                        aggregate.clone(),
                        Value::F64(i as f64),
                    );
                    let on_r = bound_settler2(
                        allsettled_rejected_idx,
                        aggregate.clone(),
                        Value::F64(i as f64),
                    );
                    add_reaction(&p, on_f, on_r, make_promise("pending", Value::Undefined));
                }
            }
            aggregate
        }),
    );

    vm.register_host_fn(
        "ecma:promise",
        "any",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let Some(inputs) = promise_combinator_inputs(ctx, args, "Promise.any") else {
                return Value::Undefined;
            };
            let n = inputs.len();
            let aggregate = pending_promise_with_id(ctx);
            let errors = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                Value::Undefined;
                n
            ]))));
            if let Value::Object(ref obj) = aggregate {
                let mut o = obj.lock().unwrap();
                o.properties.insert("__any_errors".into(), errors);
                o.properties
                    .insert("__any_remaining".into(), Value::F64(n as f64));
            }
            if n == 0 {
                settle_and_drain(
                    ctx,
                    &[aggregate.clone(), aggregate_error(vec![])],
                    "rejected",
                );
                return aggregate;
            }
            for (i, input) in inputs.into_iter().enumerate() {
                let p = promise_resolve_for_combinator(ctx, input);
                let (state, value) = read_promise_state(&p);
                if state == "fulfilled" {
                    settle_and_drain(ctx, &[aggregate.clone(), value], "fulfilled");
                    return aggregate;
                } else if state == "rejected" {
                    if let Some(error) = any_record_rejection(&aggregate, i, value) {
                        settle_and_drain(ctx, &[aggregate.clone(), error], "rejected");
                    }
                } else {
                    let on_f = bound_settler(any_fulfilled_idx, aggregate.clone());
                    let on_r =
                        bound_settler2(any_rejected_idx, aggregate.clone(), Value::F64(i as f64));
                    add_reaction(&p, on_f, on_r, make_promise("pending", Value::Undefined));
                }
            }
            aggregate
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

    // queueMicrotask(callback) — HTML global exposing the ECMA microtask queue
    // (§9.5 HostEnqueuePromiseJob checkpoint). Enqueues the callback to run
    // after the current synchronous run-to-completion, alongside promise jobs.
    vm.register_host_fn(
        "ecma:promise",
        "queueMicrotask",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let cb = args.first().cloned().unwrap_or(Value::Undefined);
            if is_callable(&cb) {
                ctx.queue_microtask(cb, Value::Undefined);
            }
            Value::Undefined
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
        "fulfilled" | "rejected" => {
            let result_promise = pending_promise_with_id(ctx);
            queue_promise_reaction(
                ctx,
                result_promise.clone(),
                &state,
                on_fulfilled,
                on_rejected,
                value,
            );
            result_promise
        }
        // Pending: register a reaction to fire when the promise settles.
        _ => {
            let result_promise = pending_promise_with_id(ctx);
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
    match state.as_str() {
        "fulfilled" | "rejected" => {
            let result_promise = pending_promise_with_id(ctx);
            let reaction = finally_reaction(result_promise.clone(), state.as_str(), on_finally);
            ctx.queue_microtask(reaction, value);
            result_promise
        }
        // Pending: register a reaction so onFinally runs (and can override with a
        // throw) once the promise settles — mirrors then_impl's pending branch.
        // The forwarders make run_reaction invoke onFinally and preserve the
        // original settlement unless onFinally throws (§27.2.5.3).
        _ => {
            let result_promise = pending_promise_with_id(ctx);
            let on_f = finalizer_forwarder(on_finally.clone(), "fulfilled");
            let on_r = finalizer_forwarder(on_finally, "rejected");
            add_reaction(&p, on_f, on_r, result_promise.clone());
            result_promise
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

fn pending_promise_with_id(ctx: &mut HostContext) -> Value {
    let promise = make_promise("pending", Value::Undefined);
    let id = ctx.next_promise_id();
    if let Value::Object(ref obj) = promise {
        obj.lock()
            .unwrap()
            .properties
            .insert("__id".into(), Value::F64(id as f64));
    }
    promise
}

fn promise_reaction(
    result_promise: Value,
    state: &str,
    on_fulfilled: Value,
    on_rejected: Value,
) -> Value {
    let idx = *PROMISE_REACTION_HOST_IDX
        .get()
        .expect("ecma:promise::__reaction registered before Promise reactions");
    let mut obj = Object::new();
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
            result_promise,
            Value::String(Arc::from(state)),
            on_fulfilled,
            on_rejected,
        ])))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn finally_reaction(result_promise: Value, state: &str, on_finally: Value) -> Value {
    promise_reaction(
        result_promise,
        state,
        finalizer_forwarder(on_finally.clone(), state),
        finalizer_forwarder(on_finally, state),
    )
}

fn finalizer_forwarder(on_finally: Value, state: &str) -> Value {
    let idx = *PROMISE_REACTION_HOST_IDX
        .get()
        .expect("ecma:promise::__reaction registered before Promise reactions");
    let mut obj = Object::new();
    obj.properties
        .insert("__promise_finally".into(), on_finally);
    obj.properties.insert(
        "__promise_finally_state".into(),
        Value::String(Arc::from(state)),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn queue_promise_reaction(
    ctx: &mut HostContext,
    result_promise: Value,
    state: &str,
    on_fulfilled: Value,
    on_rejected: Value,
    value: Value,
) {
    let reaction = promise_reaction(result_promise, state, on_fulfilled, on_rejected);
    ctx.queue_microtask(reaction, value);
}

fn run_reaction(
    ctx: &mut HostContext,
    result_promise: Value,
    state: &str,
    on_fulfilled: Value,
    on_rejected: Value,
    value: Value,
) {
    let cb = if state == "fulfilled" {
        on_fulfilled
    } else {
        on_rejected
    };

    if let Value::Object(obj) = &cb {
        let maybe_finally = {
            let o = obj.lock().unwrap();
            o.properties.get("__promise_finally").cloned()
        };
        if let Some(on_finally) = maybe_finally {
            match ctx.try_invoke(&on_finally, &[]) {
                Err(exc) => mutate_promise_state(ctx, &result_promise, "rejected", exc),
                Ok(ret) if is_promise(&ret) || get_then_method(&ret).is_some() => {
                    // §27.2.5.3: onFinally's return value is normally ignored,
                    // but a returned thenable is awaited — if it REJECTS the
                    // chain rejects with that reason; if it fulfills the original
                    // settlement is preserved.
                    let temp = pending_promise_with_id(ctx);
                    resolve_promise_with_value(ctx, &temp, ret);
                    let preserve = PRESERVE_IDX
                        .get()
                        .map(|&i| {
                            bound_settler3(
                                i,
                                result_promise.clone(),
                                Value::String(Arc::from(state)),
                                value.clone(),
                            )
                        })
                        .unwrap_or(Value::Undefined);
                    let reject_fwd = SETTLE_REJECTED_IDX
                        .get()
                        .map(|&i| bound_settler(i, result_promise.clone()))
                        .unwrap_or(Value::Undefined);
                    let (ts, tv) = read_promise_state(&temp);
                    match ts.as_str() {
                        "fulfilled" => ctx.queue_microtask(preserve, tv),
                        "rejected" => ctx.queue_microtask(reject_fwd, tv),
                        _ => {
                            add_reaction(&temp, preserve, reject_fwd, pending_promise_with_id(ctx))
                        }
                    }
                }
                Ok(_) => mutate_promise_state(ctx, &result_promise, state, value),
            }
            return;
        }
    }

    // Synthetic Promise-internal handlers.
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
            mutate_promise_state(ctx, &result_promise, "fulfilled", result);
            return;
        }
        if let Some(ret) = o.properties.get("__catch_return").cloned() {
            drop(o);
            mutate_promise_state(ctx, &result_promise, "fulfilled", ret);
            return;
        }
    }

    if !is_callable(&cb) {
        // No handler for this state — pass the settlement through unchanged.
        mutate_promise_state(ctx, &result_promise, state, value);
        return;
    }

    match ctx.try_invoke(&cb, &[value]) {
        // A throw in the handler rejects the derived promise.
        Err(exc) => mutate_promise_state(ctx, &result_promise, "rejected", exc),
        // Otherwise resolve the derived promise WITH the handler's return value,
        // assimilating a returned promise/thenable (adopt its eventual state).
        Ok(result) => resolve_promise_with_value(ctx, &result_promise, result),
    }
}

/// ECMA-262 §27.2.1.3.2 Promise Resolve Functions — settle `promise` with
/// `value`, adopting a returned native promise or raw thenable's *eventual*
/// state (register a reaction when pending) rather than fulfilling with the
/// object itself. This is what makes a handler that returns a still-pending,
/// later-rejecting promise correctly reject the chain.
fn resolve_promise_with_value(ctx: &mut HostContext, promise: &Value, value: Value) {
    if let Some(previous) = read_thenable_resolution(promise) {
        if same_object(&previous, &value) {
            let err = crate::ecma::error::new_error(
                ctx,
                "TypeError",
                "Chaining cycle detected for promise",
            );
            mutate_promise_state(ctx, promise, "rejected", err);
            return;
        }
    }

    // §27.2.1.3.2 step 6: resolving a promise with ITSELF is a TypeError
    // rejection ("Chaining cycle detected").
    if same_object(promise, &value) {
        let err =
            crate::ecma::error::new_error(ctx, "TypeError", "Chaining cycle detected for promise");
        mutate_promise_state(ctx, promise, "rejected", err);
        return;
    }
    if is_promise(&value) {
        let (state, inner) = read_promise_state(&value);
        match state.as_str() {
            "fulfilled" => resolve_promise_with_value(ctx, promise, inner),
            "rejected" => mutate_promise_state(ctx, promise, "rejected", inner),
            _ => {
                // Pending: forward `value`'s settlement onto `promise` —
                // fulfillment re-resolves (the settled value could itself
                // be a thenable), rejection forwards raw.
                if let (Some(&fi), Some(&ri)) = (RESOLVE_IDX.get(), SETTLE_REJECTED_IDX.get()) {
                    let on_f = bound_settler(fi, promise.clone());
                    let on_r = bound_settler(ri, promise.clone());
                    add_reaction(&value, on_f, on_r, pending_promise_with_id(ctx));
                }
            }
        }
        return;
    }
    // §27.2.1.3.2 steps 8–13: ONE GetV(resolution, "then"); a throwing
    // getter rejects; a callable `then` queues a NewPromiseResolveThenableJob
    // (§27.2.2.2 — runs as a MICROTASK with this = thenable).
    match get_then(ctx, &value) {
        Err(exc) => mutate_promise_state(ctx, promise, "rejected", exc),
        Ok(Some(then_fn)) => {
            remember_thenable_resolution(promise, &value);
            if let Some(&job_idx) = THENABLE_JOB_IDX.get() {
                let job = bound_settler3(job_idx, promise.clone(), value, then_fn);
                ctx.queue_microtask(job, Value::Undefined);
            }
        }
        Ok(None) => mutate_promise_state(ctx, promise, "fulfilled", value),
    }
}

fn same_object(a: &Value, b: &Value) -> bool {
    matches!((a, b), (Value::Object(left), Value::Object(right)) if Arc::ptr_eq(left, right))
}

fn remember_thenable_resolution(promise: &Value, thenable: &Value) {
    if let Value::Object(promise_obj) = promise {
        promise_obj
            .lock()
            .unwrap()
            .properties
            .insert("__resolving_thenable".into(), thenable.clone());
    }
}

fn read_thenable_resolution(promise: &Value) -> Option<Value> {
    let Value::Object(promise_obj) = promise else {
        return None;
    };
    promise_obj
        .lock()
        .unwrap()
        .properties
        .get("__resolving_thenable")
        .cloned()
}

/// §27.2.1.3.2 steps 8–9 — GetV(resolution, "then"), exactly once,
/// honoring an accessor `then` (`__get_then` convention); a throwing
/// getter surfaces as Err so the caller rejects with the thrown value.
fn get_then(ctx: &mut HostContext, val: &Value) -> Result<Option<Value>, Value> {
    let Value::Object(obj) = val else {
        return Ok(None);
    };
    let getter = { obj.lock().unwrap().properties.get("__get_then").cloned() };
    if let Some(g) = getter {
        if is_callable(&g) {
            let arity = match &g {
                Value::Object(go) => match &go.lock().unwrap().kind {
                    ObjectKind::Function(f) => f.arity,
                    _ => 0,
                },
                _ => 0,
            };
            let saved_this = ctx.current_js_this();
            ctx.set_js_this(val.clone());
            let outcome = if arity >= 1 {
                ctx.try_invoke(&g, &[val.clone()])
            } else {
                ctx.try_invoke(&g, &[])
            };
            ctx.set_js_this(saved_this);
            let f = outcome?;
            return Ok(if is_callable(&f) { Some(f) } else { None });
        }
    }
    let then_fn = { obj.lock().unwrap().properties.get("then").cloned() };
    Ok(then_fn.filter(is_callable))
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

    // Queue each reaction as a PromiseJob microtask.
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

        queue_promise_reaction(
            ctx,
            result_promise,
            state,
            on_fulfilled,
            on_rejected,
            value.clone(),
        );
    }
}

/// Overwrite a promise's state/value in-place and wake up any suspended fiber.
fn mutate_promise_state(ctx: &mut HostContext, promise: &Value, state: &str, value: Value) {
    let mut reactions: Vec<Value> = vec![];
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
            o.properties.remove("__resolving_thenable");
            o.properties
                .insert("__state".into(), Value::String(Arc::from(state)));
            o.properties.insert("__value".into(), value.clone());
            if let Some(Value::Object(arr)) = o.properties.remove("__pending_reactions") {
                let mut a = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = a.kind {
                    reactions = std::mem::take(elems);
                }
            }
            o.properties.get("__id").map(|v| v.as_f64() as u64)
        } else {
            None
        }
    } else {
        None
    };
    if let Some(id) = promise_id {
        if state == "rejected" {
            ctx.reject_promise(id, value.clone());
        } else {
            ctx.resolve_promise(id, value.clone());
        }
    }
    for reaction in reactions {
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
        queue_promise_reaction(
            ctx,
            result_promise,
            state,
            on_fulfilled,
            on_rejected,
            value.clone(),
        );
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────

fn promise_combinator_inputs(
    ctx: &mut HostContext,
    args: &[Value],
    name: &str,
) -> Option<Vec<Value>> {
    match args.first() {
        Some(Value::Object(arr)) => match &arr.lock().unwrap().kind {
            ObjectKind::Array(elems) => Some(elems.clone()),
            _ => {
                ctx.throw_value(crate::ecma::error::new_error(
                    ctx,
                    "TypeError",
                    &format!("{name} argument is not iterable"),
                ));
                None
            }
        },
        _ => {
            ctx.throw_value(crate::ecma::error::new_error(
                ctx,
                "TypeError",
                &format!("{name} argument is not iterable"),
            ));
            None
        }
    }
}

fn promise_resolve_for_combinator(ctx: &mut HostContext, value: Value) -> Value {
    if is_promise(&value) {
        value
    } else {
        let promise = pending_promise_with_id(ctx);
        resolve_promise_with_value(ctx, &promise, value);
        promise
    }
}

fn settled_descriptor(state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    if state == "rejected" {
        obj.properties
            .insert("status".into(), Value::String(Arc::from("rejected")));
        obj.properties.insert("reason".into(), value);
    } else {
        obj.properties
            .insert("status".into(), Value::String(Arc::from("fulfilled")));
        obj.properties.insert("value".into(), value);
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn aggregate_error(errors: Vec<Value>) -> Value {
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
    Value::Object(Arc::new(Mutex::new(agg)))
}

fn any_record_rejection(aggregate: &Value, index: usize, reason: Value) -> Option<Value> {
    let Value::Object(obj) = aggregate else {
        return None;
    };
    let mut o = obj.lock().unwrap();
    let errors = o.properties.get("__any_errors").cloned();
    if let Some(Value::Object(ref arr)) = errors {
        if let ObjectKind::Array(ref mut elems) = arr.lock().unwrap().kind {
            if index < elems.len() {
                elems[index] = reason;
            }
        }
    }
    let remaining = o
        .properties
        .get("__any_remaining")
        .map(|v| v.as_f64() as i64)
        .unwrap_or(0)
        - 1;
    o.properties
        .insert("__any_remaining".into(), Value::F64(remaining as f64));
    if remaining <= 0 {
        if let Some(Value::Object(arr)) = errors {
            let values = match &arr.lock().unwrap().kind {
                ObjectKind::Array(elems) => elems.clone(),
                _ => vec![],
            };
            Some(aggregate_error(values))
        } else {
            Some(aggregate_error(vec![]))
        }
    } else {
        None
    }
}

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

pub(crate) fn make_promise(state: &str, value: Value) -> Value {
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

/// Host-fn ref bound to three args (e.g. `__preserve`'s `[promise, state, value]`).
fn bound_settler3(idx: usize, a: Value, b: Value, c: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![a, b, c])))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Host-fn ref bound to two args (e.g. Promise.all's `[aggregate, index]`).
fn bound_settler2(idx: usize, a: Value, b: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![a, b])))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Record a fulfilled `Promise.all` element. Returns `Some(results)` once the
/// last outstanding element settles (so the caller settles the aggregate).
fn aggregate_record_element(aggregate: &Value, index: usize, value: Value) -> Option<Value> {
    let Value::Object(obj) = aggregate else {
        return None;
    };
    let mut o = obj.lock().unwrap();
    let results = o.properties.get("__all_results").cloned();
    if let Some(Value::Object(ref arr)) = results {
        if let ObjectKind::Array(ref mut elems) = arr.lock().unwrap().kind {
            if index < elems.len() {
                elems[index] = value;
            }
        }
    }
    let remaining = o
        .properties
        .get("__all_remaining")
        .map(|v| v.as_f64() as i64)
        .unwrap_or(0)
        - 1;
    o.properties
        .insert("__all_remaining".into(), Value::F64(remaining as f64));
    if remaining <= 0 { results } else { None }
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
