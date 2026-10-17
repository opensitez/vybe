use grameo::{CompileOptions, program::Program};
const BASE: &str = "entry = { SOI ~ expr ~ EOI }\nexpr = { atom ~ (plus ~ atom)* }\natom = @{ ASCII_DIGIT+ }\nplus = @{ \"+\" }\nWHITESPACE = _{ \" \" }\n";
fn enabled(source: &str) -> Result<grameo::CompiledGrammar, Vec<grameo::Diagnostic>> {
    grameo::compile_with_options(
        source,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
}
#[test]
fn malformed_known_directives_are_comments_by_default_and_located_only_when_enabled() {
    for line in [
        "optimize",
        "optimize unknown",
        "pratt",
        "pratt expr atom=atom trailing=wrong",
        "pratt expr atom=atom",
        "pratt expr atom=atom trailing=consume extra",
        "operator",
        "operator expr plus precedence=x fixity=infix_left",
        "operator expr plus precedence=65536 fixity=infix_left",
        "operator expr plus precedence=10 fixity=wrong",
        "operator expr absent precedence=10 fixity=infix_left",
        "pratt absent atom=atom trailing=consume",
        "pratt expr atom=absent trailing=consume",
        "pratt expr atom=expr trailing=consume",
    ] {
        let source = format!("// @grameo {line}\n{BASE}");
        let plain = grameo::compile(&source).unwrap();
        let default = grameo::compile_with_options(&source, CompileOptions::default()).unwrap();
        assert_eq!(plain.syntax(), default.syntax());
        assert_eq!(plain.optimization_selection(), Default::default());
        assert!(plain.source_pratt(1).is_none());
        grameo::compile_with_ast_bindings(&source).unwrap();
        let e = enabled(&source).unwrap_err();
        assert!(!e.is_empty());
        assert_eq!(e[0].span.start, 0);
        assert_eq!(
            &source[e[0].span.start..e[0].span.end],
            format!("// @grameo {line}")
        );
    }
}
#[test]
fn annotations_only_come_from_actual_line_comments() {
    let source = "entry = { \"// @grameo optimize unknown\" }\n/* // @grameo optimize unknown */\n/* nested /* // @grameo pratt broken */ comment */\n// @grameo optimizer unknown\n// @grameo optimizeX unknown\n";
    let g = enabled(source).unwrap();
    assert_eq!(g.optimization_selection(), Default::default());
}
#[test]
fn enabled_pratt_and_native_metadata_equal_explicit_bindings() {
    let source = format!(
        "// @grameo optimize native_choices native_rule_heads\n// @grameo pratt expr atom=atom trailing=consume\n// @grameo operator expr plus precedence=10 fixity=infix_left\n{BASE}"
    );
    let actual = enabled(&source).unwrap();
    let mut expected = grameo::compile(&source).unwrap();
    expected
        .bind_pratt(
            "expr",
            "atom",
            &[grameo::islands::Operator {
                rule: "plus",
                precedence: 10,
                fixity: grameo::pratt::Fixity::InfixLeft,
            }],
            grameo::islands::TrailingTrivia::Consume,
        )
        .unwrap();
    for id in 0..expected.syntax().rules.len() {
        assert_eq!(
            format!("{:?}", actual.source_pratt(id)),
            format!("{:?}", expected.source_pratt(id))
        );
    }
    for input in ["1", "1+2", "1 + 2 + 3", "1+", "+1", "", "1 x"] {
        assert_eq!(
            format!("{:?}", actual.recognize("entry", input)),
            format!("{:?}", expected.recognize("entry", input))
        );
    }
}
#[test]
fn matcher_selection_can_enable_choices_and_heads_independently() {
    for flag in ["native_choices", "native_rule_heads"] {
        let source = format!("// @grameo optimize {flag}\n{BASE}");
        let g = enabled(&source).unwrap();
        let selection = g.optimization_selection();
        assert_eq!(selection.native_choices, flag == "native_choices");
        assert_eq!(selection.native_rule_heads, flag == "native_rule_heads");
    }
}
#[test]
fn ast_annotations_remain_an_independent_explicit_option() {
    let source = format!(
        "// @grameo ast atom = Integer(text: source, span: rule)\n// @grameo optimize native_rule_heads\n{BASE}"
    );
    let ast = grameo::compile_with_options(
        &source,
        CompileOptions {
            ast_bindings: true,
            optimizations: false,
        },
    )
    .unwrap();
    assert_eq!(ast.syntax().ast_bindings.len(), 1);
    assert_eq!(ast.optimization_selection(), Default::default());
    let optimized = enabled(&source).unwrap();
    assert!(optimized.syntax().ast_bindings.is_empty());
    let both = grameo::compile_with_options(
        &source,
        CompileOptions {
            ast_bindings: true,
            optimizations: true,
        },
    )
    .unwrap();
    assert_eq!(both.syntax().ast_bindings.len(), 1);
    assert!(both.optimization_selection().native_rule_heads);
}
#[test]
fn duplicate_and_unbound_operator_contracts_fail_only_when_enabled() {
    for metadata in [
        "// @grameo pratt expr atom=atom trailing=consume\n// @grameo pratt expr atom=atom trailing=preserve\n",
        "// @grameo operator expr plus precedence=10 fixity=infix_left\n",
        "// @grameo pratt expr atom=atom trailing=consume\n// @grameo operator expr plus precedence=10 fixity=infix_left\n// @grameo operator expr plus precedence=20 fixity=infix_right\n",
    ] {
        let s = format!("{metadata}{BASE}");
        grameo::compile(&s).unwrap();
        assert!(enabled(&s).is_err());
    }
}
