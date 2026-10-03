#[path = "../../generated_tests/tests/support/ast_compare.rs"]
mod ast_compare;
#[path = "../../generated_tests/tests/support/lua_samples.rs"]
mod samples;
mod native {
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
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/lua_native_walker.rs"));
    }
    mod normalize {
        include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
    }
}
mod reference {
    use pest_derive::Parser;
    #[derive(Parser)]
    #[grammar = "../../../languages/lua/src/grammar.pest"]
    struct LuaParser;
    #[allow(
        clippy::collapsible_if,
        clippy::while_let_on_iterator,
        clippy::clone_on_copy,
        clippy::manual_strip
    )]
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/lua_raw_walker.rs"));
    }
    mod normalize {
        include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
    }
}
#[test]
fn real_lua_walker_raw_ast_is_structurally_identical_for_native_and_reference_pairs() {
    for source in samples::VALID {
        let a =
            reference::walker::parse(source).unwrap_or_else(|e| panic!("reference {source}: {e}"));
        let b = native::walker::parse(source).unwrap_or_else(|e| panic!("native {source}: {e}"));
        ast_compare::module(&a, &b);
    }
    let large = (0..200)
        .map(|n| format!("local x{n} = {n} + f(1,2*3)\n"))
        .collect::<String>();
    ast_compare::module(
        &reference::walker::parse(&large).unwrap(),
        &native::walker::parse(&large).unwrap(),
    );
    for source in samples::INVALID {
        assert!(reference::walker::parse(source).is_err());
        assert!(native::walker::parse(source).is_err());
    }
}
