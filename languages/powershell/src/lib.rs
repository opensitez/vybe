//! PowerShell language plugin crate.
//!
//! This crate is registered like other frontends through the shared plugin
//! registry. Parsing and AST lowering are the only language-specific pieces.

// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod emitter;
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
    // The profile inherits its platform constants through `type_scopes`, so the
    // platform has to be in the registry before the TOML is parsed. `register`
    // is idempotent, and this is the one call site guaranteed to run first.
    vybe_platform_dotnet::register();
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
        // ONE arm: `powershell.add`. PowerShell's `+` types its result from
        // the LEFT operand (array append / string concat / arithmetic) and no
        // shared primitive expresses that — see `emitter/operators.rs` for the
        // three lowerings that were tried first.
        emit_dispatch: Some(emitter::dispatch::dispatch),
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
