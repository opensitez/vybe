pub mod emitter;
pub mod normalize_class;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct RubyParser;

/// Parse Ruby source into the common AST.
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
        name: "ruby",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
}

/// This crate as a [`vybe_bytecode::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "ruby"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}
