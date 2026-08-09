//! The ecma platform as a `vybe_runtime::Plugin` — one plugin, same type as
//! every other. Its `init` registers what this JS-runtime provider offers:
//! the `ecma:*` host functions, the `wasm:js-*` primitive builtins the runtime
//! builds on, and `ecma:date` (Clock-gated). Its `finalize` wires the ECMA
//! global objects (constructor↔prototype links, `globalThis`, `Intl`), which
//! must run AFTER every plugin's host fns are registered.

use vybe_runtime::capabilities::Capability;

/// The ecma platform plugin.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "ecma"
    }

    // No `reset`, and the split is worth stating because this platform holds
    // both kinds of state:
    //
    // The shared prototypes stay in `OnceLock` statics, and must: they are
    // primed before the boot snapshot and live in the tracked heap, so
    // `reset_to` already restores their contents (a program that assigns
    // `Array.prototype.map` has it rolled back). They are boot state — putting
    // them in the resource store would DELETE them on every reset.
    //
    // The Error constructor cache is the opposite case — built lazily while a
    // program runs, hence allocated after the snapshot and gutted by the very
    // reset that is supposed to restore it — so it is VM-owned storage
    // (`vybe_runtime::resources`) and `reset_to` drops it.

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        // ecma:date reads the system clock — gated under Clock.
        let clock = fw.granted(Capability::Clock);
        if let Some(vm) = fw.vm.as_deref_mut() {
            // ECMA-262 host functions.
            crate::register(vm);
            // The §9.5 job scheduler — policy moves to the layer whose spec
            // defines it; the VM keeps only mechanism (JSPI suspend/resume).
            vm.set_scheduler(std::sync::Arc::new(crate::scheduler::EcmaScheduler));
            // wasm:js-* CG proposals (js-string / js-primitive builtins) — the
            // WASM-level JS primitives the ecma runtime is built on.
            vybe_runtime::js_builtins::register(vm);
            if clock {
                crate::date::register(vm);
            }
        }
    }

    fn finalize(&self, fw: &mut vybe_runtime::Framework<'_>) {
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

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
