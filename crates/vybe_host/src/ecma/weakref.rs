//! # `ecma:weakref` and `ecma:finalization-registry` host handlers
//!
//! Native Rust impls of `WeakRef.*` per ECMA-262 §26.1 and
//! `FinalizationRegistry.*` per §26.2.
//!
//! ## Weak reference caveat (same as `ecma:weakmap`)
//!
//! WASM GC MVP doesn't yet expose weak references (it's a Post-MVP
//! proposal). The MVP-compatible stand-in stores a *strong* reference
//! to the target on the WeakRef object and always returns it from
//! `deref()`. Spec-compatible at the API surface (constructor +
//! `deref`); the semantic difference is that the target is never
//! reclaimed while the WeakRef is alive (vs. `wr.deref() === undefined`
//! after collection on v8/SpiderMonkey).
//!
//! Phase B4 (the same upgrade earmarked for `WeakMap`) swaps the
//! strong refs for `std::sync::Weak<Mutex<Object>>` once the VM tracks
//! object lifetime independently from the user-visible Value graph.
//! Until then:
//!   * Functional correctness is preserved for code that just round-
//!     trips the target.
//!   * Memory-pressure-driven semantics (caches that empty after GC,
//!     subscriber lists that auto-prune) retain more memory than on
//!     a v8/SpiderMonkey runtime.
//!
//! ## FinalizationRegistry semantics
//!
//! Per spec, `FinalizationRegistry(cleanupCallback)` invokes the
//! callback when a previously-`register()`-ed target gets garbage-
//! collected. With strong refs as the backing model, targets are
//! never collected while the registry holds them, so the callback
//! never fires. We expose the API (`register` / `unregister`) so
//! code that *constructs* a FinalizationRegistry doesn't fault, and
//! `unregister` removes the entry the same way the spec does.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

const MODULE_WEAKREF: &str = "ecma:weakref";
const MODULE_REGISTRY: &str = "ecma:finalization-registry";

const WEAKREF_TAG: &str = "__vybe_js_weakref";
const WEAKREF_TARGET_PROP: &str = "__vybe_wr_target";

const REGISTRY_TAG: &str = "__vybe_js_finreg";
const REGISTRY_CALLBACK_PROP: &str = "__vybe_fr_cb";
const REGISTRY_ENTRIES_PROP: &str = "__vybe_fr_entries";

// ── Constructors ──────────────────────────────────────────────────

fn new_weakref(target: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("WeakRef")));
    obj.properties.insert(WEAKREF_TAG.into(), Value::I32(1));
    obj.properties.insert(WEAKREF_TARGET_PROP.into(), target);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn new_finalization_registry(callback: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(),
        Value::String(Arc::from("FinalizationRegistry")));
    obj.properties.insert(REGISTRY_TAG.into(), Value::I32(1));
    obj.properties.insert(REGISTRY_CALLBACK_PROP.into(), callback);
    // Entries: Array of `[target, heldValue, unregisterToken]` tuples.
    obj.properties.insert(
        REGISTRY_ENTRIES_PROP.into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))),
    );
    Value::Object(Arc::new(Mutex::new(obj)))
}

// ── Type guards ───────────────────────────────────────────────────

fn is_weakref(arg: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(obj) = arg {
        let o = obj.lock().unwrap();
        if o.properties.get(WEAKREF_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn is_registry(arg: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(obj) = arg {
        let o = obj.lock().unwrap();
        if o.properties.get(REGISTRY_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

// ── Registration ──────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    // ── WeakRef (§26.1) ──────────────────────────────────────────

    // `new WeakRef(target)` — wraps target. ECMA-262 §26.1.1.1: target must
    // be an Object (or non-registered Symbol since ES2024). Throw TypeError
    // for primitives, null, and undefined.
    vm.register_host_fn(MODULE_WEAKREF, "new", Box::new(|ctx, args| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        match &target {
            Value::Object(_) | Value::Symbol(_) => {}
            _ => {
                ctx.throw_value(crate::ecma::error::new_error(
                    "TypeError",
                    "WeakRef target must be an object",
                ));
                return Value::Undefined;
            }
        }
        new_weakref(target)
    }));

    // `wr.deref()` — returns the wrapped target, or `undefined` after
    // collection. Strong-ref stand-in always returns the target.
    vm.register_host_fn(MODULE_WEAKREF, "deref", Box::new(|_ctx, args| {
        let Some(wr) = args.first().and_then(is_weakref) else {
            return Value::Undefined;
        };
        let lock = wr.lock().unwrap();
        lock.properties.get(WEAKREF_TARGET_PROP).cloned().unwrap_or(Value::Undefined)
    }));

    // ── FinalizationRegistry (§26.2) ─────────────────────────────

    // `new FinalizationRegistry(cleanupCallback)` — stores the
    // callback. With strong-ref backing the callback is never invoked,
    // but the constructor still has to stash it so user code that
    // checks `.constructor` / type doesn't fault.
    vm.register_host_fn(MODULE_REGISTRY, "new", Box::new(|_ctx, args| {
        let callback = args.first().cloned().unwrap_or(Value::Undefined);
        new_finalization_registry(callback)
    }));

    // `registry.register(target, heldValue, [unregisterToken])` —
    // appends `[target, heldValue, token?]` to the entries Array.
    vm.register_host_fn(MODULE_REGISTRY, "register", Box::new(|_ctx, args| {
        let Some(registry) = args.first().and_then(is_registry) else {
            return Value::Undefined;
        };
        let target = args.get(1).cloned().unwrap_or(Value::Undefined);
        let held = args.get(2).cloned().unwrap_or(Value::Undefined);
        let token = args.get(3).cloned().unwrap_or(Value::Undefined);

        let entry = Value::Object(Arc::new(Mutex::new(
            Object::new_array(vec![target, held, token])
        )));

        let lock = registry.lock().unwrap();
        if let Some(Value::Object(entries)) = lock.properties.get(REGISTRY_ENTRIES_PROP) {
            let entries = entries.clone();
            drop(lock);
            if let ObjectKind::Array(ref mut items) = entries.lock().unwrap().kind {
                items.push(entry);
            }
        }
        Value::Undefined
    }));

    // `registry.registerWithToken(target, heldValue, unregisterToken)` — alias.
    vm.register_host_fn(MODULE_REGISTRY, "registerWithToken", Box::new(|_ctx, args| {
        let Some(registry) = args.first().and_then(is_registry) else {
            return Value::Undefined;
        };
        let target = args.get(1).cloned().unwrap_or(Value::Undefined);
        let held = args.get(2).cloned().unwrap_or(Value::Undefined);
        let token = args.get(3).cloned().unwrap_or(Value::Undefined);
        let entry = Value::Object(Arc::new(Mutex::new(
            Object::new_array(vec![target, held, token])
        )));
        let lock = registry.lock().unwrap();
        if let Some(Value::Object(entries)) = lock.properties.get(REGISTRY_ENTRIES_PROP) {
            let entries = entries.clone();
            drop(lock);
            if let ObjectKind::Array(ref mut items) = entries.lock().unwrap().kind {
                items.push(entry);
            }
        }
        Value::Undefined
    }));

    // `registry.pendingCleanupCount()` — always 0 in sync tests.
    vm.register_host_fn(MODULE_REGISTRY, "pendingCleanupCount", Box::new(|_ctx, _args| {
        Value::I32(0)
    }));

    // `registry.unregister(unregisterToken)` — removes every entry
    // whose `unregisterToken` SameValue-matches the argument. Returns
    // `true` if at least one entry was removed (per spec §26.2.3.3).
    vm.register_host_fn(MODULE_REGISTRY, "unregister", Box::new(|_ctx, args| {
        let Some(registry) = args.first().and_then(is_registry) else {
            return Value::Bool(false);
        };
        let token = args.get(1).cloned().unwrap_or(Value::Undefined);
        if matches!(token, Value::Undefined | Value::Null) {
            return Value::Bool(false);
        }

        let lock = registry.lock().unwrap();
        let Some(Value::Object(entries)) = lock.properties.get(REGISTRY_ENTRIES_PROP) else {
            return Value::Bool(false);
        };
        let entries = entries.clone();
        drop(lock);

        let mut removed = false;
        if let ObjectKind::Array(ref mut items) = entries.lock().unwrap().kind {
            let before = items.len();
            items.retain(|entry| {
                if let Value::Object(t) = entry {
                    let t = t.lock().unwrap();
                    if let ObjectKind::Array(ref tuple) = t.kind {
                        if let Some(stored_token) = tuple.get(2) {
                            if same_value(stored_token, &token) {
                                return false;
                            }
                        }
                    }
                }
                true
            });
            removed = items.len() < before;
        }
        Value::Bool(removed)
    }));
}

// ── SameValue (ECMA-262 §7.2.11) ─────────────────────────────────
//
// FinalizationRegistry.unregister uses SameValue to match tokens.
// We use object-identity comparison for Object values (Arc pointer
// equality), which is what the spec mandates for non-primitive types.

fn same_value(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Object(x), Value::Object(y)) => Arc::ptr_eq(x, y),
        (Value::String(x), Value::String(y)) => x.as_ref() == y.as_ref(),
        (Value::I32(x), Value::I32(y)) => x == y,
        (Value::I64(x), Value::I64(y)) => x == y,
        (Value::F64(x), Value::F64(y)) => {
            // SameValue semantics: NaN === NaN, +0 !== -0.
            if x.is_nan() && y.is_nan() {
                true
            } else if *x == 0.0 && *y == 0.0 {
                x.to_bits() == y.to_bits()
            } else {
                x == y
            }
        }
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::Null, Value::Null) => true,
        (Value::Undefined, Value::Undefined) => true,
        _ => false,
    }
}

// ── Instance method dispatch ───────────────────────────────────────
//
// Called from `ecma:value::dispatch_plain_object` when the receiver's
// `__type` is "WeakRef" or "FinalizationRegistry" and the method isn't
// found as an own property (which it never is — host fns live in the
// type registry, not on the object itself).

pub fn dispatch_weakref_method(obj: Arc<Mutex<Object>>, method: &str, args: &[Value]) -> Option<Value> {
    let wr = Value::Object(obj);
    match method {
        "deref" => {
            let Some(arc) = is_weakref(&wr) else { return None };
            let lock = arc.lock().unwrap();
            Some(lock.properties.get(WEAKREF_TARGET_PROP).cloned().unwrap_or(Value::Undefined))
        }
        _ => None,
    }
}

pub fn dispatch_registry_method(obj: Arc<Mutex<Object>>, method: &str, args: &[Value]) -> Option<Value> {
    let reg = Value::Object(obj.clone());
    match method {
        "register" => {
            let Some(registry) = is_registry(&reg) else { return None };
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let held = args.get(1).cloned().unwrap_or(Value::Undefined);
            let token = args.get(2).cloned().unwrap_or(Value::Undefined);
            let entry = Value::Object(Arc::new(Mutex::new(
                Object::new_array(vec![target, held, token])
            )));
            let lock = registry.lock().unwrap();
            if let Some(Value::Object(entries)) = lock.properties.get(REGISTRY_ENTRIES_PROP) {
                let entries = entries.clone();
                drop(lock);
                if let ObjectKind::Array(ref mut items) = entries.lock().unwrap().kind {
                    items.push(entry);
                }
            }
            Some(Value::Undefined)
        }
        "unregister" => {
            let Some(registry) = is_registry(&reg) else { return None };
            let token = args.first().cloned().unwrap_or(Value::Undefined);
            if matches!(token, Value::Undefined | Value::Null) {
                return Some(Value::Bool(false));
            }
            let lock = registry.lock().unwrap();
            let Some(Value::Object(entries)) = lock.properties.get(REGISTRY_ENTRIES_PROP) else {
                return Some(Value::Bool(false));
            };
            let entries = entries.clone();
            drop(lock);
            let mut removed = false;
            if let ObjectKind::Array(ref mut items) = entries.lock().unwrap().kind {
                let before = items.len();
                items.retain(|entry| {
                    if let Value::Object(t) = entry {
                        let t = t.lock().unwrap();
                        if let ObjectKind::Array(ref tuple) = t.kind {
                            if let Some(stored_token) = tuple.get(2) {
                                if same_value(stored_token, &token) {
                                    return false;
                                }
                            }
                        }
                    }
                    true
                });
                removed = items.len() < before;
            }
            Some(Value::Bool(removed))
        }
        _ => None,
    }
}
