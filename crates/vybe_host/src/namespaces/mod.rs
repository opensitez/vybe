//! Namespace objects — sets up VM globals as nested objects
//! that any language can traverse via struct_get chains.
//!
//! Each sub-module registers one domain (math, io, forms, etc.).

mod array;
mod console;
mod data;
mod ecma_globals;
mod environment;
mod forms;
mod gui;
mod intl;
mod io;
mod json;
mod math;
mod net;
mod string;
mod threading;
mod types;
mod vb_globals;
mod web_globals;

use std::sync::Arc;
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value};

/// Register all namespace objects on the VM.
/// Call this AFTER registering all host functions.
pub fn setup_namespaces(vm: &mut VM) {
    math::register(vm);
    console::register(vm);
    string::register(vm);
    array::register(vm);
    io::register(vm);
    environment::register(vm);
    forms::register(vm);
    gui::register(vm);
    net::register(vm);
    json::register(vm);
    types::register(vm);
    threading::register(vm);
    data::register(vm);
    intl::register(vm);
    vb_globals::register(vm);
    // New ECMA-262 globals (Symbol, Reflect, Atomics, BigInt, Iterator,
    // Math accumulator extensions) and the proper §19.3 globalThis
    // singleton — must run AFTER the legacy ad-hoc globalThis stub
    // below so its singleton overwrites the placeholder.
    ecma_globals::register(vm);
    // WHATWG / W3C web platform globals: crypto, URL, TextEncoder/Decoder,
    // fetch + Headers/Request/Response, setTimeout / setInterval / etc.
    web_globals::register(vm);

    // Polyfill: now that all host fns are registered (including any test
    // overrides), point each `__vybe_*` global at its native equivalent so
    // calls compiled as `global_get __vybe_pow + call_ref` execute the
    // optimized native version instead of the bundled stdlib bytecode.
    crate::modules::override_stdlib_globals_with_host_fns(vm);
}

// ---- Shared helpers used by all sub-modules ----

/// Ensure a namespace object exists at the given dotted path.
/// Creates intermediate objects as needed. Returns the leaf object.
///
/// Stores the root global under BOTH the original case and lowercase.
/// Case-sensitive language lookups (JS `Intl.X` / `Math.PI`) hit the
/// original-case key directly via GLOBAL_GET; case-insensitive language
/// lookups (VB `system.console.writeline`) hit the lowercase key — same
/// underlying Object, so writes through either name are visible from
/// the other.
pub(crate) fn ensure_namespace(vm: &mut VM, path: &[&str]) -> Value {
    if path.is_empty() {
        return Value::Null;
    }
    let root_orig = path[0].to_string();
    let root_lc = root_orig.to_lowercase();
    let root = if let Some(existing) = vm
        .globals
        .get(&root_orig)
        .or_else(|| vm.globals.get(&root_lc))
    {
        existing.clone()
    } else {
        let obj = Value::Object(vybe_bytecode::heap::alloc(Object::new()));
        vm.globals.insert(root_orig.clone(), obj.clone());
        if root_lc != root_orig {
            vm.globals.insert(root_lc, obj.clone());
        }
        obj
    };
    let mut current = root;
    for &segment in &path[1..] {
        let orig = segment.to_string();
        let key_lc = orig.to_lowercase();
        let next = if let Value::Object(ref obj) = current {
            let lock = obj.lock().unwrap();
            lock.properties
                .get(&orig)
                .or_else(|| lock.properties.get(&key_lc))
                .cloned()
        } else {
            None
        };
        if let Some(existing) = next {
            current = existing;
        } else {
            let new_obj = Value::Object(vybe_bytecode::heap::alloc(Object::new()));
            if let Value::Object(ref obj) = current {
                let mut o = obj.lock().unwrap();
                o.properties.insert(orig.clone(), new_obj.clone());
                if key_lc != orig {
                    o.properties.insert(key_lc, new_obj.clone());
                }
            }
            current = new_obj;
        }
    }
    current
}

/// Set a property on a namespace object. Stores under BOTH the
/// original-case key (for JS / case-sensitive consumers) and the
/// lowercase form (for VB / case-insensitive consumers). Both keys
/// reference the same Value, so semantics stay consistent.
pub(crate) fn set_prop(ns: &Value, name: &str, value: Value) {
    if let Value::Object(obj) = ns {
        let lc = name.to_lowercase();
        let mut o = obj.lock().unwrap();
        o.properties.insert(name.to_string(), value.clone());
        if lc != name {
            o.properties.insert(lc, value);
        }
    }
}

/// Wire a shared prototype singleton's `constructor` exactly once
/// (first-writer-wins) and return the canonical constructor that's now on it.
///
/// The prototype objects are process-global `OnceLock`s, but each VM that runs
/// `ecma_globals::register` would otherwise re-point `proto.constructor` at its
/// own fresh constructor — so under parallel VMs (e.g. the test harness) a value
/// stamped by one VM and a `Ctor` read by another could disagree. Pinning the
/// constructor on first write makes `x.constructor === Ctor` hold across VMs;
/// the constructor's host-fn refs use deterministic registry indices, so the
/// first VM's constructor is valid in every VM.
pub(crate) fn set_constructor_once(proto: &Value, ctor: Value) -> Value {
    if let Value::Object(obj) = proto {
        let mut o = obj.lock().unwrap();
        if let Some(existing) = o.properties.get("constructor") {
            return existing.clone();
        }
        o.properties.insert("constructor".to_string(), ctor.clone());
    }
    ctor
}

/// Create a HostFunction Value referencing a registered host function.
pub(crate) fn host_fn_ref(vm: &VM, module: &str, name: &str) -> Value {
    if let Some(&idx) = vm
        .host_registry
        .get(&(module.to_string(), name.to_string()))
    {
        let mut obj = Object::new();
        obj.properties
            .insert("__host_module".into(), Value::String(Arc::from(module)));
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.properties.insert(
            "__proto__".into(),
            crate::ecma::function::shared_function_prototype(),
        );
        obj.properties
            .insert("name".into(), Value::String(Arc::from(name)));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_bytecode::heap::alloc(obj))
    } else {
        Value::Null
    }
}

// `receiver_host_fn_ref` / `bound_host_fn_ref` moved into `vybe_platform_ecma`
// (they stamp the shared ecma function prototype, so they belong with ecma).
// `receiver_host_fn_ref` is re-exported here so existing `crate::namespaces::…`
// call sites keep resolving; `bound_host_fn_ref` has no remaining vybe_host
// consumer (its only user moved into the ecma crate).
pub(crate) use crate::ecma::receiver_host_fn_ref;
