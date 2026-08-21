// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod designer_codegen;
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct VbParser;

/// Parse VB source into the common AST.
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

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Platforms this language needs. A language already links them (see
    // Cargo.toml), so registering them here means the compiler never has to —
    // and it works from ANY host, including language test binaries that never
    // construct vybex. See flexclassplan.md.
    vybe_platform_dotnet::register();

    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "vb",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
        expand_source: None,
    });
    vybe_platform_dotnet::winforms::form_modules::register(
        vybe_platform_dotnet::winforms::form_modules::FormModuleLanguage {
            name: "vb",
            generate_designer_code: designer_codegen::generate_designer_code,
            generate_user_code_stub: designer_codegen::generate_user_code_stub,
        },
    );
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "vb"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
