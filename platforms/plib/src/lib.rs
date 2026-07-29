//! Pascal library platform package.
//!
//! The compiler-side code generation surface lives under [`emitter`].

pub mod emitter;

/// Register this platform with the shared registry.
///
/// The plugin entry point: the host calls this after linking — or, once
/// platforms ship as dylibs, after `dlopen`. `vybe_compiler` must NOT depend on
/// this crate; it reaches the platform through `vybe_runtime::registry`
/// function pointers, exactly as it already does for languages.
pub fn register() {
    // Mount the namespace tree HERE, at plugin-registration time. Previously
    // `vybe_compiler` called this from `resolver.rs` via a hard Cargo
    // dependency, which is what prevented this crate from ever being a dylib.
    emitter::tree_register::register_namespace_tree();
    vybe_runtime::registry::register_platform(vybe_runtime::registry::PlatformDef {
        name: "plib",
        emit_dispatch: None,
        register_tree: Some(crate::emitter::tree_register::register_namespace_tree),
        namespace_constants: None,
        component_descriptor: None,
        is_descriptor_class: None,
        numeric_format_helper: None,
        read_binary_module: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — the SAME single plugin trait
/// languages use. `init` registers the platform's compiler-facing surface, and
/// this is the dylib entry point. Mirrors `vybe_language_*::Plugin` exactly, so
/// there is one plugin concept in the tree, not two.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "plib"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
