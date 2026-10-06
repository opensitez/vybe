#[path = "../../generated_tests/tests/support/ast_compare.rs"]
mod ast_compare;
#[path = "../../generated_tests/tests/support/lua_ast.rs"]
mod direct;
#[path = "../../generated_tests/tests/support/lua_expression_samples.rs"]
mod expression_samples;
#[path = "../../generated_tests/tests/support/lua_module_samples.rs"]
mod module_samples;
#[path = "../../generated_tests/tests/support/lua_samples.rs"]
mod samples;

#[test]
fn pratt_direct_lua_ast_matches_reference_with_complete_spans_and_module_policies() {
    use grameo::{MatchOptions, engine::build_program};
    use vybe_parser_generated_tests::lua_pratt::Parser;
    for source in expression_samples::cases() {
        let mut builder = direct::Ast::with_pratt(&source);
        let matched = build_program(
            &Parser,
            "expr",
            &source,
            MatchOptions::default(),
            &mut builder,
        );
        match (
            matched,
            reference::walker::expression_for_conformance(&source),
        ) {
            (Ok(matched), Ok((reference, consumed))) => {
                assert_eq!(matched.consumed, consumed, "{source}");
                ast_compare::expression(&builder.into_expression(), &reference);
            }
            (Err(_), Err(_)) => assert!(builder.is_empty(), "{source}"),
            (actual, reference) => panic!("{source}: {actual:?}; {reference:?}"),
        }
    }
    for source in module_samples::cases() {
        let mut builder = direct::Ast::with_pratt(&source);
        let matched = build_program(
            &Parser,
            "chunk",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        assert_eq!(matched.consumed, source.len());
        ast_compare::module(
            &builder.into_module(),
            &reference::walker::parse(&source).unwrap(),
        );
    }
}

#[test]
fn direct_lua_modules_match_reference_including_policies_and_spans() {
    use grameo::{MatchOptions, engine::build_program};
    use vybe_parser_generated_tests::lua::Parser;
    for source in module_samples::cases() {
        let mut builder = direct::Ast::new(&source);
        let matched = build_program(
            &Parser,
            "chunk",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        assert_eq!(matched.consumed, source.len());
        let actual = builder.into_module();
        let reference = reference::walker::parse(&source).unwrap();
        ast_compare::module(&actual, &reference);
    }
}
mod native {
    use vybe_parser_generated_tests::lua::{Parser, Rule};
    struct LuaParser;
    impl LuaParser {
        fn parse(
            rule: Rule,
            source: &str,
        ) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
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

#[test]
fn direct_lua_expression_ast_matches_black_box_reference_walker_including_spans() {
    use grameo::{MatchOptions, builder::Builder, engine::build_program};
    use vybe_parser_generated_tests::lua::Parser;
    for source in expression_samples::cases() {
        let mut builder = direct::Ast::new(&source);
        let result = build_program(
            &Parser,
            "expr",
            &source,
            MatchOptions::default(),
            &mut builder,
        );
        match (
            result,
            reference::walker::expression_for_conformance(&source),
        ) {
            (Ok(result), Ok((reference, consumed))) => {
                assert_eq!(result.consumed, consumed, "{source}");
                let actual = builder.into_expression();
                assert!(
                    std::panic::catch_unwind(|| ast_compare::expression(&actual, &reference))
                        .is_ok(),
                    "{source}: direct={actual:?}; reference={reference:?}"
                );
            }
            (Err(_), Err(_)) => {
                assert_eq!(builder.checkpoint(), 0);
                assert!(builder.is_empty());
            }
            (actual, reference) => panic!("{source}: direct={actual:?}; reference={reference:?}"),
        }
    }
}
