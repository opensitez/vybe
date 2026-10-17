// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod core_classes;
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod walker;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct PythonParser;
impl PythonParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}

/// Parse Python source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse_with_direct_literals(source)
}

/// Explicit PEG comparison frontend. Registration uses the owned hybrid path.
pub fn parse_with_peg(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Deferred Pratt expression recipes lower common AST in the current scope.
/// Rich operands and statements retain their original captures and walker.
/// This is the same frontend used by the registered [`parse`] entry point.
pub fn parse_with_direct_ast(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse_with_direct_literals(source)
}

/// Compatibility entry point for the original literal-list experiment.
pub fn parse_with_direct_literals(source: &str) -> Result<vybe_ast::Module, String> {
    parse_with_direct_ast(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "python",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(emitter::tree_register::register_namespace_tree),
        expand_source: None,
    });
    vybe_runtime::registry::register_hooks(
        "python",
        vybe_runtime::registry::LanguageHooks {
            ..Default::default()
        },
    );
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "python"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);

/// Explicit eager comparison frontend for the optimized grammar.
pub fn parse_with_eager_ast(source: &str) -> Result<vybe_ast::Module, String> { walker::parse_with_backend(source, true, false) }
