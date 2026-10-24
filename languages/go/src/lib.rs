// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod adapters;
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod tree_register;
pub mod walker;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub mod direct_ast;
pub(crate) struct GoParser;
impl GoParser {
    /// Ordinary captures for a rule of any source, including fragments.
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        direct_ast::parse_reference(rule, source)
    }
    /// A Go source file, with the expressions the grammar hooks construct
    /// directly; see [`direct_ast`].
    pub(crate) fn parse_program(source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        direct_ast::parse_owned(Rule::program, source)
    }
}

/// Parse Go source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, profile) = direct_ast::profile_construction(|| walker::parse(source));
    direct_ast::clear_owned();
    // An attached debugger report lists the construction counts per parse.
    drop(vybe_runtime::debugger::DebugPhase::current_lazy(|| format!(
        "go expressions owned={} handed_off={} walked={}",
        profile.owned_regions, profile.handoffs, profile.walked)));
    module
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Exact-parser defaults preserve an explicit caller's capability selection,
    // including disabling shared helpers. Other capabilities remain unchanged.
    vybe_compiler::frontend_preparation::register_default(parse,
        vybe_compiler::frontend_preparation::Capabilities {
            helper_definitions: emitter::dispatch::helper_definitions(),
            ..Default::default()
        });
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "go",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "go"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
