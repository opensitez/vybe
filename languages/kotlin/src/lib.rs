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
    // The JDK is a PLATFORM Kotlin consumes — same relationship csharp/vb have
    // with dotnet. Kotlin declares NO `java.*` names of its own; it declares
    // tree data in its profile and the common resolver does the rest.
    vybe_platform_jvm::register();
    // The declarations point at `common:java.*` emit targets, which still live
    // in Java's emitter and dispatch through its LanguageDef. That coupling is
    // real and temporary — javakotlinmigration.md Phases 3-6 move the adapters
    // into the platform, and this line goes away with the last one.
    vybe_language_java::register();
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
