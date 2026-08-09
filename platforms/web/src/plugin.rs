//! The web platform as a `vybe_runtime::Plugin` — one plugin, same type as all
//! the others. `init` registers the `web:*` host functions (WHATWG / W3C:
//! crypto, URL, TextEncoder, fetch, dom-parser). Always-on (pure computation).

/// The web platform plugin.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "web"
    }

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        if let Some(vm) = fw.vm.as_deref_mut() {
            crate::register(vm);
        }
    }

    fn finalize(&self, fw: &mut vybe_runtime::Framework<'_>) {
        // The web-platform TypeRegistry vtables (TextEncoder/Decoder,
        // URLSearchParams, Response, DOM node hierarchy) + DOM type-id
        // stamping — registered via the `register_type` primitive after every
        // plugin's host fns exist.
        crate::builtin_types::register_types(fw);
    }

    // No `reset`. Everything this platform holds on a program's behalf — the
    // DOM listener table, the ambient document — is VM-owned storage
    // (`vybe_runtime::resources`), so `reset_to` drops it without this plugin
    // taking part. That is the whole point of the store: the listener table
    // used to be a process-global static with a hand-written `reset_listeners`,
    // and `reset_active_document`'s only caller was a pascal test helper — a
    // per-test helper cannot be the mechanism, it fixes the one caller that
    // remembers to call it and leaves every other embedder broken.
    //
    // Queued timer and animation callbacks are not here either: they are
    // `DeferredSource`s, and `reset_to` clears every registered source's queue
    // through `clear_pending`. That is what a `setTimeout` callback outliving
    // its program used to defeat — it stayed queued and was drained by the NEXT
    // program, against chunk indices that had since been reused, which is how
    // one tenant's code came to run under another's closure.
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
