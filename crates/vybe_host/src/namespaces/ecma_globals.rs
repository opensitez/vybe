//! New ECMA-262 global constructors and namespaces.
//!
//! Wires up `Symbol`, `Reflect`, `Atomics`, `BigInt`, `Iterator`, plus
//! `Math` Stage-3 accumulator extensions (`minOf`, `maxOf`,
//! `sumPrecise`) and the proper `globalThis` singleton — all so JS code
//! can do `Symbol.iterator`, `Reflect.apply(...)`, `Atomics.add(...)`,
//! `BigInt(123n)`, etc., the same way real JS runtimes expose them.
//!
//! Each `Foo` global is an Object whose properties are HostFunction refs
//! pointing at the registered `ecma:foo:*` host fns. Static methods
//! (`Symbol.for`, `Reflect.apply`) become properties; constructors
//! (`Symbol`, `BigInt`) double as the namespace object's own
//! invocation target via the `__call` convention.

use super::*;

pub fn register(vm: &mut VM) {
    // ── Symbol ─────────────────────────────────────────────────────
    // `Symbol` is callable as `Symbol(desc)` AND has static properties
    // (`Symbol.for`, `Symbol.iterator`, etc.). Build a Value that is
    // both a HostFunction (kind=HostFunction(idx)) and carries the
    // static properties as Object properties.
    let sym = host_fn_with_props(vm, "ecma:symbol", "Symbol");
    set_prop(&sym, "for",            host_fn_ref(vm, "ecma:symbol", "for"));
    set_prop(&sym, "keyFor",         host_fn_ref(vm, "ecma:symbol", "keyFor"));
    // Well-knowns — materialized as Symbol primitives directly. The
    // sentinel strings are the same that `ecma:symbol` exposes via
    // its 0-arg getters (kept in sync; both paths share the convention
    // that well-known symbols stringify to "@@<name>").
    for name in &[
        "iterator", "asyncIterator", "toPrimitive", "hasInstance",
        "toStringTag", "isConcatSpreadable", "unscopables", "match",
        "matchAll", "replace", "search", "split", "species", "dispose",
        "asyncDispose",
    ] {
        let sentinel = format!("@@{}", name);
        set_prop(&sym, name, Value::Symbol(Arc::from(sentinel.as_str())));
    }
    vm.globals.insert("Symbol".to_string(), sym);
    vm.globals.insert("symbol".to_string(), vm.globals.get("Symbol").cloned().unwrap_or(Value::Null));

    // ── Reflect ────────────────────────────────────────────────────
    let reflect = ensure_namespace(vm, &["Reflect"]);
    for name in &[
        "apply", "construct", "get", "set", "has", "deleteProperty",
        "ownKeys", "getOwnPropertyDescriptor", "defineProperty",
        "getPrototypeOf", "setPrototypeOf", "isExtensible",
        "preventExtensions",
    ] {
        set_prop(&reflect, name, host_fn_ref(vm, "ecma:reflect", name));
    }

    // ── Atomics ────────────────────────────────────────────────────
    let atomics = ensure_namespace(vm, &["Atomics"]);
    for name in &[
        "add", "sub", "and", "or", "xor", "exchange", "compareExchange",
        "load", "store", "isLockFree", "wait", "notify",
    ] {
        set_prop(&atomics, name, host_fn_ref(vm, "ecma:atomics", name));
    }

    // ── BigInt ─────────────────────────────────────────────────────
    // Same callable-with-static-properties shape as Symbol.
    let bigint = host_fn_with_props(vm, "ecma:bigint", "BigInt");
    set_prop(&bigint, "asIntN",  host_fn_ref(vm, "ecma:bigint", "asIntN"));
    set_prop(&bigint, "asUintN", host_fn_ref(vm, "ecma:bigint", "asUintN"));
    vm.globals.insert("BigInt".to_string(), bigint);
    vm.globals.insert("bigint".to_string(), vm.globals.get("BigInt").cloned().unwrap_or(Value::Null));

    // ── Iterator (Stage-3 helpers) ─────────────────────────────────
    let iter = ensure_namespace(vm, &["Iterator"]);
    set_prop(&iter, "from",  host_fn_ref(vm, "ecma:iterator", "from"));
    set_prop(&iter, "range", host_fn_ref(vm, "ecma:iterator", "range"));

    // ── Math Stage-3 accumulators ──────────────────────────────────
    // Math global already created by namespaces/math.rs; just add the new
    // properties so they show up alongside the existing ones.
    let math = ensure_namespace(vm, &["Math"]);
    set_prop(&math, "minOf",       host_fn_ref(vm, "ecma:math", "minOf"));
    set_prop(&math, "maxOf",       host_fn_ref(vm, "ecma:math", "maxOf"));
    set_prop(&math, "sumPrecise",  host_fn_ref(vm, "ecma:math", "sumPrecise"));

    // ── globalThis — proper §19.3.1 singleton ──────────────────────
    // Pulls the shared process-global Object that `ecma:globalThis.get`
    // also returns, so identity holds across both access patterns.
    vm.globals.insert("globalThis".to_string(), crate::ecma::global_this::shared_singleton());
}

/// Build a Value that is callable as a host function AND can carry
/// static properties — like real JS where `Symbol` and `BigInt` are
/// both callable AND have static methods (`Symbol.for`, `BigInt.asIntN`).
fn host_fn_with_props(vm: &VM, module: &str, name: &str) -> Value {
    host_fn_ref(vm, module, name)
}
