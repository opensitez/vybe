#[path = "support/lua_samples.rs"]
mod samples;
use vybe_parser_generated_tests::lua::{Parser, Rule};
struct LuaParser;
impl LuaParser {
    fn parse(
        rule: Rule,
        source: &str,
    ) -> Result<vybe_parser::compat::Pairs<'_, Rule>, vybe_parser::ParseError> {
        Parser.parse_pairs(rule, source)
    }
}
#[allow(
    clippy::collapsible_if,
    clippy::while_let_on_iterator,
    clippy::clone_on_copy,
    clippy::manual_strip
)]
mod walker {
    include!(concat!(env!("OUT_DIR"), "/lua_native_walker.rs"));
}
mod normalize {
    include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
}

#[test]
fn actual_lua_walker_runs_on_native_pairs_without_compiler_or_runtime_dependencies() {
    for source in samples::VALID {
        let module = walker::parse(source).unwrap_or_else(|e| panic!("{source}: {e}"));
        assert_eq!(module.language, vybe_ast::Lang::Lua);
        assert_eq!(module.name, "main");
    }
    for source in samples::INVALID {
        assert!(walker::parse(source).is_err(), "{source}");
    }
}
