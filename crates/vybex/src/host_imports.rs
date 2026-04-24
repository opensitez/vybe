//! Runtime wiring for ESM host-module imports.
//!
//! Given the `HostImportMetadata` produced by the compiler, install the
//! runtime bindings the user's code will read by name:
//!
//! * `import { log } from "wasi:cli"` — install `log` as a global bound
//!   to the host function reference so `const f = log; f("hi")` works
//!   (the compiler already emits a direct `CALL_IMPORT` for `log(...)`).
//! * `import * as cli from "wasi:cli"` — synthesize a namespace object
//!   exposing every host function registered under `wasi:cli` as a
//!   property, and install it as a global under the alias.
//!
//! Both forms read the same source of truth — `vm.host_registry` and
//! `vm.func_table` — which is populated before the script runs by the
//! `register_host_fn` calls in `vybe_host::register_*`.
//!
//! The compiler itself never emits dispatch code against these globals
//! for direct calls: those become `CALL_IMPORT`. This module only
//! matters for read-as-value and namespace access.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{Value, VM};
use vybe_bytecode::value::{Object, ObjectKind};

use crate::compiler::HostImportMetadata;

/// Install named + wildcard ESM host-module imports as VM globals.
pub fn install(vm: &mut VM, meta: &HostImportMetadata) {
    for n in &meta.named {
        let Some(val) = host_fn_value(vm, &n.module, &n.func) else {
            // Unresolved named import — leave it to `setup_execution` to
            // surface as an `Unresolved import` error if the code actually
            // calls it.
            continue;
        };
        vm.globals.insert(n.local.clone(), val);
    }

    for w in &meta.wildcard {
        let ns = build_namespace(vm, &w.module);
        vm.globals.insert(w.alias.clone(), ns);
    }
}

/// Look up a host function by `(module, name)` and return a reusable
/// `Value::Object(HostFunction(idx))` reference from the func table.
fn host_fn_value(vm: &VM, module: &str, func: &str) -> Option<Value> {
    let idx = *vm.host_registry.get(&(module.to_string(), func.to_string()))?;
    vm.func_table.get(idx).cloned()
}

/// Build the Module Namespace Object for `import * as ns from "<module>"`
/// per ECMA-262 §10.4.6.
///
/// Own keys are the module's exports in sorted order (spec: the [[OwnPropertyKeys]]
/// internal method returns exports sorted by ECMA standard string
/// ordering). `@@toStringTag` is represented indirectly — the VM's
/// `Display` impl on `ObjectKind::ModuleNamespace` renders the value as
/// `"[object Module]"`, matching `Object.prototype.toString.call(ns)`.
///
/// The object is marked `ObjectKind::ModuleNamespace` so downstream
/// code (e.g. `typeof`) can recognize it. Mutations to property values
/// are blocked conceptually by the kind tag, though enforcement lives
/// in the VM's property-set path (Phase 5b; today the object is
/// constructed once and never handed to mutable code paths).
fn build_namespace(vm: &VM, module: &str) -> Value {
    // Collect + sort exports so ECMA-262 §10.4.6 [[OwnPropertyKeys]]
    // returns them in a deterministic, spec-aligned order. Consumers
    // that care about iteration order (`Object.keys(ns)`,
    // `Reflect.ownKeys(ns)`, `for ... in ns`) see the same order every
    // time.
    let mut exports: Vec<(String, Value)> = vm.host_registry.iter()
        .filter_map(|((m, name), &idx)| {
            if m != module { return None; }
            vm.func_table.get(idx).cloned().map(|v| (name.clone(), v))
        })
        .collect();
    exports.sort_by(|a, b| a.0.cmp(&b.0));

    let mut obj = Object::new();
    obj.kind = ObjectKind::ModuleNamespace;
    for (name, val) in exports {
        obj.properties.insert(name, val);
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}
