//! The web platform as a `vybe_bytecode::Plugin` — one plugin, same type as all
//! the others. `init` registers the `web:*` host functions (WHATWG / W3C:
//! crypto, URL, TextEncoder, fetch, dom-parser). Always-on (pure computation).

/// The web platform plugin.
pub struct Plugin;

impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "web"
    }

    fn init(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        if let Some(vm) = fw.vm.as_deref_mut() {
            crate::register(vm);
        }
    }

    fn finalize(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        // The web-platform TypeRegistry vtables (TextEncoder/Decoder,
        // URLSearchParams, Response, DOM node hierarchy) + DOM type-id
        // stamping — registered via the `register_type` primitive after every
        // plugin's host fns exist.
        crate::builtin_types::register_types(fw);
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_bytecode::register_plugin!(Plugin);
