//! PowerShell language plugin crate.
//!
//! This crate is registered like other frontends through the shared plugin
//! registry. Parsing and AST lowering are the only language-specific pieces.

// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod normalize_class;
pub mod protocol;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub struct PowerShellParser;

/// Parse PowerShell source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    static DOTNET_CONSTANTS: std::sync::Once = std::sync::Once::new();
    DOTNET_CONSTANTS.call_once(|| {
        let mappings = vybe_platform_dotnet::emitter::namespace_constant_mappings()
            .iter()
            .map(|(name, value)| (name.to_string(), *value))
            .collect();
        vybe_runtime::profile::register_dotnet_namespace_constants(mappings);
    });
    include_str!("profile")
}

/// Register this language with the shared plugin registry.
pub fn register() {
    // Platforms this language needs.
    vybe_platform_dotnet::register();

    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "powershell",
        parse,
        profile_source,
        emit_dispatch: None,
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
        expand_source: None,
    });
}

/// Dylib entry point — registers this language.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "powershell"
    }

    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
