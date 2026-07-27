// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod designer_codegen;
pub mod emitter;
pub mod forms;
pub mod normalize_class;
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
    static DOTNET_CONSTANTS: std::sync::Once = std::sync::Once::new();
    DOTNET_CONSTANTS.call_once(|| {
        let mappings = vybe_platform_dotnet::emitter::namespace_constant_mappings()
            .iter()
            .map(|(name, value)| (name.to_string(), *value))
            .collect();
        vybe_bytecode::profile::register_dotnet_namespace_constants(mappings);
    });
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Platforms this language needs. A language already links them (see
    // Cargo.toml), so registering them here means the compiler never has to —
    // and it works from ANY host, including language test binaries that never
    // construct vybex. See flexclassplan.md.
    vybe_platform_dotnet::register();

    vybe_bytecode::registry::register_language(vybe_bytecode::registry::LanguageDef {
        name: "vb",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
    vybe_platform_dotnet::winforms::form_modules::register(
        vybe_platform_dotnet::winforms::form_modules::FormModuleLanguage {
            name: "vb",
            load_designer: forms::load_designer,
            save_designer: forms::save_designer,
            generate_designer_code: designer_codegen::generate_designer_code,
            generate_user_code_stub: designer_codegen::generate_user_code_stub,
        },
    );
}

/// This crate as a [`vybe_bytecode::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "vb"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_bytecode::register_plugin!(Plugin);
