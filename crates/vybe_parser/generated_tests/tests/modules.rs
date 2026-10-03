use vybe_parser::{
    MatchOptions,
    engine::parse_program,
    modules::{self, Import, Module},
    program::Program,
};
use vybe_parser_generated_tests::modular;

#[test]
fn generated_linked_modules_preserve_scope_aliases_captures_and_failures() {
    let linked = modules::compile(&[
        Module {
            name: "outer",
            source: include_str!("modular_outer.grammar"),
            exports: &["program"],
            imports: &[Import {
                local: "imported",
                module: "inner",
                rule: "word",
            }],
        },
        Module {
            name: "inner",
            source: include_str!("modular_inner.grammar"),
            exports: &["word"],
            imports: &[],
        },
    ])
    .unwrap();
    let native = linked.grammar();
    assert_eq!(
        modular::Parser.rule_id("outer::program"),
        native.rule_id("outer::program")
    );
    for input in [
        "a x_y b", "ax_yb", "a x y b", "a_x_y_b", "a x_y b!", "a xx b",
    ] {
        match (
            native.parse_source("outer::program", input),
            parse_program(
                &modular::Parser,
                "outer::program",
                input,
                MatchOptions::default(),
            ),
        ) {
            (Ok(a), Ok(b)) => {
                assert_eq!(a.consumed(), b.consumed());
                let flatten = |tree: &vybe_parser::ParseTree<'_, '_>| {
                    tree.pairs()
                        .flat_map(|pair| pair.into_inner())
                        .map(|pair| (pair.rule_name().to_owned(), pair.as_span()))
                        .collect::<Vec<_>>()
                };
                assert_eq!(flatten(&a), flatten(&b));
            }
            (Err(a), Err(b)) => assert_eq!(a, b),
            (a, b) => panic!("{input:?}: {a:?} vs {b:?}"),
        }
    }
    assert!(
        modular::Parser
            .parse_pairs(modular::Rule::m0_program, "a x_y b")
            .is_ok()
    );
}
