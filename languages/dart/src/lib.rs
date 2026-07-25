pub mod emitter;
pub mod normalize_class;
pub mod tree_register;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct DartParser;

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
        name: "dart",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
    });
    // Dart records compare by value while Lists compare by reference, so the
    // equality fallback deep-compares only tagged tuples (see vybe_emitter).
    vybe_bytecode::registry::register_hooks(
        "dart",
        vybe_bytecode::registry::LanguageHooks {
            value_eq: Some(vybe_emitter::tuples::emit_tuple_value_eq),
            ..Default::default()
        },
    );
}

/// This crate as a [`vybe_bytecode::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "dart"
    }
    fn init(&self, _fw: &mut vybe_bytecode::Framework<'_>) {
        register();
    }
}
