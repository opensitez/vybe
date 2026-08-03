//! Runtime wiring for ESM host-module imports.
//!
//! Given the `HostImportMetadata` produced by the compiler, install the
//! runtime bindings the user's code will read by name:
//!
//! * `import { log } from "wasi:cli"` — install `log` as a global bound
//!   to the host function reference so `const f = log; f("hi")` works
//!   (the compiler already emits a direct `CALL_IMPORT` for `log(...)`).
//! * `import * as cli from "wasi:cli"` — synthesize a namespace object
//!   exposing every resolvable host export registered under `wasi:cli`
//!   as a property, and install it as a global under the alias.
//!
//! Both forms read the same source of truth — `vm.modules` — so ESM
//! runtime binding installation observes the same synthetic module
//! registry as the linker and component model.
//!
//! The compiler itself never emits dispatch code against these globals
//! for direct calls: those become `CALL_IMPORT`. This module only
//! matters for read-as-value and namespace access.

use vybe_runtime::module_record::ExportEntry;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

use crate::primitives::HostImportMetadata;

/// Install named + wildcard ESM host-module imports as VM globals.
pub fn install(vm: &mut VM, meta: &HostImportMetadata) {
    for n in &meta.named {
        let Some(val) = export_value(vm, &n.module, &n.func) else {
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

/// Resolve a runtime ESM binding from the authoritative ModuleRecord view.
/// Functions reuse the VM's `func_table` marker value; immutable value
/// exports clone their registered `Value` directly.
fn export_value(vm: &VM, module: &str, name: &str) -> Option<Value> {
    resolve_export_value(vm, module, name, &mut Vec::new())
}

fn resolve_export_value(
    vm: &VM,
    module: &str,
    name: &str,
    visited: &mut Vec<(String, String)>,
) -> Option<Value> {
    let key = (module.to_string(), name.to_string());
    if visited.contains(&key) {
        return None;
    }
    visited.push(key);

    let record = vm.modules.get(module)?;
    match record.exports.get(name)? {
        ExportEntry::Function { idx } => vm.func_table.get(*idx).cloned(),
        ExportEntry::Value(value) => Some(value.clone()),
        ExportEntry::Indirect {
            from,
            name: target_name } => resolve_export_value(vm, from, target_name, visited),
        ExportEntry::Class { .. } | ExportEntry::ResourceType { .. } => None }
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
    let mut exports: Vec<(String, Value)> = vm
        .modules
        .get(module)
        .into_iter()
        .flat_map(|record| record.exports.keys())
        .filter_map(|name| export_value(vm, module, name).map(|value| (name.clone(), value)))
        .collect();
    exports.sort_by(|a, b| a.0.cmp(&b.0));

    let mut obj = Object::new();
    obj.kind = ObjectKind::ModuleNamespace;
    for (name, val) in exports {
        obj.properties.insert(name, val);
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
}
