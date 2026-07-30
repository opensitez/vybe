// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod emitter;
pub mod normalize_class;
pub mod protocol;
mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct KotlinParser;

/// Parse Kotlin source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry.
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "kotlin",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language with the shared framework.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "kotlin"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the registry.
vybe_runtime::register_plugin!(Plugin);
