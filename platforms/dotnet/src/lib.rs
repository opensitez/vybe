//! .NET-shaped platform package.
//!
//! The compiler-side code generation surface lives under [`emitter`].

pub mod emitter;
pub mod winforms;

/// Register this platform with the shared registry.
///
/// The plugin entry point: the host calls this after linking — or, once
/// platforms ship as dylibs, after `dlopen`. `vybe_compiler` must NOT depend on
/// this crate; it reaches the platform through `vybe_bytecode::registry`
/// function pointers, exactly as it already does for languages.
pub fn register() {
    // Mount the namespace tree HERE, at plugin-registration time. Previously
    // `vybe_compiler` called this from `resolver.rs` via a hard Cargo
    // dependency, which is what prevented this crate from ever being a dylib.
    emitter::tree_register::register_namespace_tree();
    vybe_bytecode::registry::register_platform(vybe_bytecode::registry::PlatformDef {
        name: "dotnet",
        emit_dispatch: Some(crate::emitter::dispatch::dispatch),
        register_tree: Some(crate::emitter::tree_register::register_namespace_tree),
        namespace_constants: Some(crate::emitter::namespace_constant_mappings),
        component_descriptor: Some(crate::emitter::dotnet_component_descriptor),
        is_descriptor_class: Some(crate::emitter::is_component_descriptor_class),
        numeric_format_helper: Some(crate::emitter::core::numeric_format::build_dotnet_numeric_format),
        read_binary_module: None,
    });
}

/// This crate as a [`vybe_bytecode::Plugin`] — the SAME single plugin trait
/// languages use. `init` registers the platform's compiler-facing surface, and
/// this is the dylib entry point. Mirrors `vybe_language_*::Plugin` exactly, so
/// there is one plugin concept in the tree, not two.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "dotnet"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_bytecode::register_plugin!(Plugin);
