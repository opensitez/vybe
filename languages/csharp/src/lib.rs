// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod forms;
pub mod normalize_class;
pub mod protocol;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct CSharpParser;

/// Parse C# source into the common AST.
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

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Platforms this language needs. A language already links them (see
    // Cargo.toml), so registering them here means the compiler never has to —
    // and it works from ANY host, including language test binaries that never
    // construct vybex. See flexclassplan.md.
    vybe_platform_dotnet::register();

    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "csharp",
        parse,
        profile_source,
        emit_dispatch: None,
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
        expand_source: None,
    });
    vybe_platform_dotnet::winforms::form_modules::register(
        vybe_platform_dotnet::winforms::form_modules::FormModuleLanguage {
            name: "csharp",
            load_designer: forms::load_designer,
            save_designer: forms::save_designer,
            generate_designer_code: forms::generate_designer_code,
            generate_user_code_stub: forms::generate_user_code_stub,
        },
    );
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "csharp"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
