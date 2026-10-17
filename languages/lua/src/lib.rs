// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
mod normalize;
pub mod normalize_class;
mod protocol;
pub mod tree_register;
pub mod walker;

#[cfg(test)]
mod conditional_tests {
    #[test]
    fn empty_then_and_elseif_bodies_keep_their_branch_roles() {
        use vybe_ast::StmtKind;
        for (source, elif_count, else_len) in [
            ("if false then else return 2 end", 0, Some(1)),
            ("if false then elseif true then return 2 else return 3 end", 1, Some(1)),
            ("if false then elseif true then else end", 1, Some(0)),
            ("if false then elseif true then end", 1, None),
        ] {
            let module = super::parse(source).unwrap();
            let statement = module.body.iter().find(|statement| matches!(statement.kind, StmtKind::If { .. })).expect("missing if statement");
            let StmtKind::If { then_body, elifs, else_body, .. } = &statement.kind else { unreachable!() };
            assert!(then_body.is_empty(), "{source}");
            assert_eq!(elifs.len(), elif_count, "{source}");
            assert_eq!(else_body.as_ref().map(Vec::len), else_len, "{source}");
        }
    }
}

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct LuaParser;
impl LuaParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}

/// Parse Lua source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "lua",
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
        "lua"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);

