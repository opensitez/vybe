// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct PythonParser;

/// Parse Python source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_bytecode::registry::register_language(vybe_bytecode::registry::LanguageDef {
        name: "python",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(emitter::tree_register::register_namespace_tree),
    });
    vybe_bytecode::registry::register_hooks(
        "python",
        vybe_bytecode::registry::LanguageHooks {
            value_eq: Some(emitter::runtime_adapter::emit_py_value_eq),
            ..Default::default()
        },
    );
}

/// This crate as a [`vybe_bytecode::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "python"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_bytecode::register_plugin!(Plugin);
