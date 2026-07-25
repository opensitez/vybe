//! The ecma platform as a `vybe_bytecode::Plugin` — one plugin, same type as
//! every other. Its `init` registers what this JS-runtime provider offers:
//! the `ecma:*` host functions, the `wasm:js-*` primitive builtins the runtime
//! builds on, and `ecma:date` (Clock-gated). Its `finalize` wires the ECMA
//! global objects (constructor↔prototype links, `globalThis`, `Intl`), which
//! must run AFTER every plugin's host fns are registered.

use vybe_bytecode::capabilities::Capability;

/// The ecma platform plugin.
pub struct Plugin;

impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "ecma"
    }

    fn init(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        // ecma:date reads the system clock — gated under Clock.
        let clock = fw.granted(Capability::Clock);
        if let Some(vm) = fw.vm.as_deref_mut() {
            // ECMA-262 host functions.
            crate::register(vm);
            // wasm:js-* CG proposals (js-string / js-primitive builtins) — the
            // WASM-level JS primitives the ecma runtime is built on.
            vybe_bytecode::js_builtins::register(vm);
            if clock {
                crate::date::register(vm);
            }
        }
    }

    fn finalize(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        // Constructor↔prototype wiring, globalThis, and Intl globals — resolves
        // host fns by registry index, so it runs after ALL plugins' inits.
        if let Some(vm) = fw.vm.as_deref_mut() {
            crate::register_globals(vm);
        }
        // The JS/Intl runtime TypeRegistry vtables (String/Map/Set/Date/RegExp/
        // Iterator/Promise/Intl.*/Object.prototype), registered via the
        // `register_type` primitive — the plugin declares its own types.
        crate::builtin_types::register_types(fw);
    }
}
