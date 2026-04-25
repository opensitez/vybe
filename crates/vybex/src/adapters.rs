//! In-language adapter modules — the Phase 6 Layer 3.
//!
//! An **Adapter module** is a JS (or future: Python / PHP / etc.) source
//! file that re-exports from Synthetic modules under a different
//! specifier. The canonical use case is Node compatibility:
//!
//! ```js
//! // node:http.js — an Adapter module
//! export { listen, createServer } from "wasi:http/server";
//! ```
//!
//! User code that writes `import { listen } from "node:http"` gets the
//! same binding as if it had imported from `wasi:http/server` directly.
//! The Linker resolves the Indirect chain at compile time, so at emit
//! time the call still becomes `CALL_IMPORT (wasi:http/server, listen)`
//! — no runtime indirection.
//!
//! Adapters live in JS (the target language's stdlib) per the
//! `feedback_layer_3_is_in_language` rule. Their sources are bundled
//! into the binary via `include_str!` and compiled + registered at VM
//! setup by [`register_all`].

use vybe_bytecode::{VM, ModuleRecord, ModuleKind, ModuleStatus, ExportEntry};

/// Register every bundled adapter module against the VM.
///
/// Order matters: adapters that re-export from other adapters must be
/// registered after their sources. Today all adapters re-export
/// exclusively from Synthetic (`wasi:*` / `wasm:js-*` / `vybe:*`)
/// modules, so order within this function is arbitrary.
///
/// Adapters are JS source files compiled at registration time. A parse
/// or link error aborts setup — adapters ship with the binary so any
/// breakage is a build error, not a runtime error.
pub fn register_all(vm: &mut VM) -> Result<(), String> {
    // `node:fs`, `node:os`, `node:path`, `node:process`,
    // `node:child_process`, `node:crypto` are now real host modules
    // under `crates/vybe_host/src/node/`. The `is_host_specifier`
    // linker path resolves their ESM imports directly. Adapters here
    // are for `node:*` modules without a real host implementation
    // yet (only `node:http` for now — `vybe:http/server.listen` has
    // a closure shape that's not trivial to expose under a new
    // module name without restructuring).
    register_js_adapter(vm, "node:http",   include_str!("adapters/node/http.js"))?;
    Ok(())
}

/// Compile a JS adapter source into a ModuleRecord and install it on
/// the VM. Today this handles pure re-export adapters — `export { X }
/// from "m"` and `export * from "m"`. Local declarations with their own
/// bodies (real translation logic) are a follow-up; those require
/// compiling the adapter's chunks into the VM's chunk list.
fn register_js_adapter(vm: &mut VM, specifier: &str, source: &str) -> Result<(), String> {
    use crate::ast::StmtKind;
    let module = crate::languages::js::parse(source)
        .map_err(|e| format!("adapter {} parse error: {}", specifier, e))?;

    let mut record = ModuleRecord {
        specifier: specifier.to_string(),
        kind: ModuleKind::Adapter,
        status: ModuleStatus::Linked,
        exports: std::collections::HashMap::new(),
        requested_modules: Vec::new(),
        capability: None,
    };

    // Walk every top-level Export statement in the adapter. Each one
    // produces ExportEntry::Indirect bindings into the adapter's
    // ModuleRecord. Star re-exports (`export * from "m"`) expand to
    // one Indirect per name in the source module.
    for stmt in &module.body {
        let StmtKind::Export { names, from, star, .. } = &stmt.kind else { continue };
        let Some(from_spec) = from else {
            // Local export — not supported yet in adapter modules.
            // This path hits if the adapter has `export function f() {...}`
            // or similar in-language translation logic. Future work.
            continue;
        };
        if *star {
            // `export * from "m"` — mirror every export from the source
            // under the same name. Per ECMA-262 §16.2.1.6.3
            // GetExportedNames, default is excluded.
            let Some(source_mod) = vm.modules.get(from_spec) else {
                return Err(format!(
                    "adapter {}: re-export source {:?} not registered",
                    specifier, from_spec
                ));
            };
            let names_in_source: Vec<String> = source_mod.exports.keys().cloned().collect();
            for name in names_in_source {
                record.exports.insert(
                    name.clone(),
                    ExportEntry::Indirect { from: from_spec.clone(), name },
                );
            }
        } else {
            for n in names {
                if n.name == "*" {
                    // `export * as n from "m"` — the whole namespace as
                    // one property. Not yet supported here; skip.
                    continue;
                }
                let local = n.alias.clone().unwrap_or_else(|| n.name.clone());
                record.exports.insert(
                    local,
                    ExportEntry::Indirect {
                        from: from_spec.clone(),
                        name: n.name.clone(),
                    },
                );
            }
        }
    }

    vm.modules.insert(specifier.to_string(), record);
    Ok(())
}
