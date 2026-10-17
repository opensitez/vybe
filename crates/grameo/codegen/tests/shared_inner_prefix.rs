use grameo::{program::Program, CompileOptions};
use grameo_codegen::{emit_with_api, ParserApi};

#[test]
fn python_prefix_directives_are_opt_in_and_generate_both_certificates() {
    let source = include_str!("../../../../languages/python/src/grammar.pest");
    let plain = grameo::compile(source).unwrap();
    let stripped = source
        .lines()
        .filter(|line| !line.contains("@grameo shared_inner_prefix"))
        .collect::<Vec<_>>()
        .join("\n")
        + "\n";
    let without = grameo::compile(&stripped).unwrap();
    assert_eq!(
        emit_with_api(&plain, ParserApi::HYBRID_CAPTURES).unwrap(),
        emit_with_api(&without, ParserApi::HYBRID_CAPTURES).unwrap()
    );
    let optimized = grameo::compile_with_options(
        source,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    let selected: Vec<_> = (0..optimized.syntax().expressions.len())
        .filter_map(|id| optimized.shared_inner_prefix(id))
        .collect();
    assert_eq!(selected.len(), 2);
    for name in ["list_inner", "subscript_item"] {
        assert!(selected
            .iter()
            .any(|s| s.target == optimized.rule_id(name).unwrap()));
    }
    let generated = emit_with_api(&optimized, ParserApi::HYBRID_CAPTURES).unwrap();
    assert_eq!(generated.matches("fn shared_inner_prefix(").count(), 1);
    assert_eq!(generated.matches("fn match_observers_are_noop(").count(), 1);
}

#[test]
fn malformed_prefix_directives_only_fail_in_optimized_mode() {
    let body = "expression={\"x\"} list={expression~clause+|(expression~(\",\"~expression)*)?~\",\"?} clause={\"for\"}";
    for directive in [
        "",
        "missing prefix=expression kind=list",
        "list prefix=missing kind=list",
        "list prefix=expression kind=unknown",
        "list prefix=expression",
        "expression prefix=expression kind=list",
    ] {
        let source = format!("// @grameo shared_inner_prefix {directive}\n{body}");
        assert!(grameo::compile(&source).is_ok());
        assert!(grameo::compile_with_options(
            &source,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            }
        )
        .is_err());
    }
}
