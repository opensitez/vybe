/// Register the SDL adapter namespace (`sdl`) and platform metadata.
pub fn register() {
    crate::tree_register::register_namespace_tree();
    vybe_runtime::registry::register_platform(vybe_runtime::registry::PlatformDef {
        name: "sdl",
        emit_dispatch: Some(crate::emitter::dispatch::dispatch),
        register_tree: Some(crate::tree_register::register_namespace_tree),
        namespace_constants: None,
        component_descriptor: None,
        is_descriptor_class: None,
        numeric_format_helper: None,
        read_binary_module: None,
    });
}

/// Register adapter as a normal `vybe_runtime::Plugin`, with no host-function
/// registrations inside this crate.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "sdl"
    }

    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
vybe_runtime::register_plugin!(Plugin);
