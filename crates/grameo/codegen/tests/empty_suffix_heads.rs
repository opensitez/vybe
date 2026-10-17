use grameo::{CompileOptions, compile_with_options};
const BODY: &str = r#"WHITESPACE=_{" "|"\t"}
COMMENT=_{"//" ~ (!"\n" ~ ANY)*}
root={word ~ (pad ~ "?" ~ word)?}
word=@{ASCII_ALPHA+}
pad=_{("\n"|";")*}
"#;
#[test]
fn annotation_is_ignored_by_default_and_codegen_is_identical() {
    let s = format!("// @grameo empty_suffix_heads root\n{BODY}");
    assert_eq!(
        grameo_codegen::generate(BODY).unwrap(),
        grameo_codegen::generate(&s).unwrap()
    );
    let _old = grameo::optimization::OptimizationSelection {
        ascii_xid: false,
        nullable_failure_heads: false,
        native_choices: false,
        native_rule_heads: false,
    };
}
#[test]
fn only_explicit_selection_emits_permission() {
    let o = CompileOptions {
        optimizations: true,
        ..Default::default()
    };
    let a = compile_with_options(BODY, o).unwrap();
    let b =
        compile_with_options(&format!("// @grameo empty_suffix_heads root\n{BODY}"), o).unwrap();
    let a = grameo_codegen::emit(&a).unwrap();
    let b = grameo_codegen::emit(&b).unwrap();
    assert!(!a.contains("fn empty_suffix_candidate"));
    assert!(b.contains("fn empty_suffix_candidate"));
    assert!(b.contains("fn empty_suffix_metadata_is_current(&self)->bool{true}"));
    assert!(b.contains("EmptySuffixHead"));
}
#[test]
fn unsupported_shapes_and_annotation_order_reject() {
    for body in [
        "root={\"a\" ~ (PUSH(\"x\") ~ \"?\")?}",
        "root={\"a\" ~ (pad ~ \"?\")?} pad={\";\"*}",
        "root={\"a\" ~ ((\"\"|\"a\") ~ \"?\")?}",
        "root={\"a\" ~ (\"?\" ~ \"b\")? ~ (\"x\")?}",
        "root={\"a\" ~ (&\"a\" ~ \"?\")?}",
        "WHITESPACE=_{ASCII_ALPHA+} root={\"a\" ~ (\"?\")?}",
    ] {
        assert!(
            compile_with_options(
                &format!("// @grameo empty_suffix_heads root\n{body}"),
                CompileOptions {
                    optimizations: true,
                    ..Default::default()
                }
            )
            .is_err(),
            "{body}"
        );
    }
    for prefix in [
        "// @grameo empty_suffix_heads root root\n",
        "// @grameo empty_suffix_heads missing\n",
        "// @grameo empty_suffix_heads root\n// @grameo pratt root atom=word trailing=consume\n",
        "// @grameo pratt root atom=word trailing=consume\n// @grameo empty_suffix_heads root\n",
    ] {
        assert!(
            compile_with_options(
                &format!("{prefix}{BODY}"),
                CompileOptions {
                    optimizations: true,
                    ..Default::default()
                }
            )
            .is_err(),
            "{prefix}"
        );
    }
}
