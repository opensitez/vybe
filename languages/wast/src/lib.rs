pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct WastParser;

/// Parse WAST/WAT source into the common AST.
/// Both .wast (full script with assertions) and .wat (module only) are handled
/// by the same grammar — WAT is a strict subset of WAST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_plugin::registry::register_language(vybe_plugin::registry::LanguagePlugin {
        name: "wast",
        parse,
        profile_source,
        emit_dispatch: None,
        normalize_class: None,
        register_tree: None,
    });
}

/// This crate as a [`vybe_plugin::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_plugin::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "wast"
    }
    fn init(&self, _fw: &mut vybe_plugin::Framework<'_>) {
        register();
    }
}
