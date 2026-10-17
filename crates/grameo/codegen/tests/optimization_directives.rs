use grameo::CompileOptions;
use grameo_codegen::{ParserApi, emit_with_api, emit_with_native_hybrid_choices};
const SOURCE: &str = "entry = { a | b }\na = { \"a\" }\nb = { \"b\" }\n";
#[test]
fn shared_prefix_is_default_inert_and_has_one_observer_admission() {
    let body = "arg={first|gen|value}\nfirst={\"!\"}\ngen={value~\"?\"}\nvalue={\"x\"}\n";
    let source = format!("// @grameo shared_prefix arg branch=gen prefix=value\n{body}");
    let plain = grameo::compile(body).unwrap();
    let ignored = grameo::compile(&source).unwrap();
    assert_eq!(emit_with_api(&plain, ParserApi::HYBRID_CAPTURES).unwrap(), emit_with_api(&ignored, ParserApi::HYBRID_CAPTURES).unwrap());
    for nullable in [false, true] {
        let input = if nullable { format!("// @grameo optimize nullable_failure_heads\n{source}") } else { source.clone() };
        let enabled = grameo::compile_with_options(&input, CompileOptions { optimizations: true, ..Default::default() }).unwrap();
        let code = emit_with_api(&enabled, ParserApi::HYBRID_CAPTURES).unwrap();
        assert!(code.contains("fn shared_prefix_choice"));
        assert_eq!(code.matches("fn match_observers_are_noop").count(), 1);
    }
}
#[test]
fn compile_options_select_native_generation_without_changing_default_or_explicit_override() {
    let plain = grameo::compile(SOURCE).unwrap();
    let defaults = emit_with_api(&plain, ParserApi::HYBRID_CAPTURES).unwrap();
    for flags in [
        "native_choices",
        "native_rule_heads",
        "native_choices native_rule_heads",
    ] {
        let source = format!("// @grameo optimize {flags}\n{SOURCE}");
        let ignored = grameo::compile(&source).unwrap();
        assert_eq!(
            emit_with_api(&ignored, ParserApi::HYBRID_CAPTURES).unwrap(),
            defaults
        );
        let enabled = grameo::compile_with_options(
            &source,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            },
        )
        .unwrap();
        let text = emit_with_api(&enabled, ParserApi::HYBRID_CAPTURES).unwrap();
        assert!(text.contains("fn native_hybrid_choices"));
        assert_eq!(
            text.contains("fn native_hybrid_choice_prefix"),
            flags.contains("native_choices")
        );
        assert_eq!(
            text.contains("fn native_hybrid_rule_head"),
            flags.contains("native_rule_heads")
        );
        assert_eq!(
            emit_with_native_hybrid_choices(&ignored, ParserApi::HYBRID_CAPTURES).unwrap(),
            emit_with_native_hybrid_choices(&enabled, ParserApi::HYBRID_CAPTURES).unwrap()
        );
    }
}
#[test]
fn annotated_pratt_tables_generate_exactly_the_explicit_binding_parser() {
    let source = "// @grameo optimize native_choices native_rule_heads\n// @grameo pratt expr atom=atom trailing=consume\n// @grameo operator expr plus precedence=10 fixity=infix_left\nexpr = { atom ~ (plus ~ atom)* }\natom = @{ ASCII_DIGIT+ }\nplus = @{ \"+\" }\n";
    let mut explicit = grameo::compile(source).unwrap();
    explicit
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
    let annotated = grameo::compile_with_options(
        source,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(
        emit_with_api(&annotated, ParserApi::HYBRID_CAPTURES).unwrap(),
        emit_with_native_hybrid_choices(&explicit, ParserApi::HYBRID_CAPTURES).unwrap()
    );
}

#[test]
fn ascii_xid_requires_its_own_explicit_compilation_selection() {
    let source = "root = { word }\nword = @{ (XID_START | \"_\") ~ XID_CONTINUE* }\n";
    let plain = grameo::compile(source).unwrap();
    let expected = emit_with_api(&plain, ParserApi::HYBRID_CAPTURES).unwrap();
    let annotation = format!("// @grameo optimize ascii_xid\n{source}");
    let ignored = grameo::compile(&annotation).unwrap();
    assert!(!ignored.optimization_selection().ascii_xid);
    assert_eq!(emit_with_api(&ignored, ParserApi::HYBRID_CAPTURES).unwrap(), expected);
    let native = grameo::compile_with_options(
        &format!("// @grameo optimize native_choices native_rule_heads\n{source}"),
        CompileOptions { optimizations: true, ..Default::default() },
    ).unwrap();
    assert!(!native.optimization_selection().ascii_xid);
    assert!(!emit_with_api(&native, ParserApi::HYBRID_CAPTURES).unwrap().contains("fn partial_atomic_fragment"));
    let opted = grameo::compile_with_options(
        &annotation, CompileOptions { optimizations: true, ..Default::default() },
    ).unwrap();
    assert!(opted.optimization_selection().ascii_xid);
    assert!(emit_with_api(&opted, ParserApi::HYBRID_CAPTURES).unwrap().contains("fn partial_atomic_fragment"));
}

#[test]
fn shared_wrapper_prefix_codegen_is_opt_in_and_static() {
    let body = "target={call|read|other}\ncall={head~suffix}\nread={head}\nhead={\"x\"}\nsuffix={\"?\"}\nother={\"!\"}\n";
    let source = format!("// @grameo shared_wrapper_prefix target branch=call fallback=read\n{body}");
    let plain=grameo::compile(body).unwrap();let ignored=grameo::compile(&source).unwrap();
    assert_eq!(emit_with_api(&plain,ParserApi::HYBRID_CAPTURES).unwrap(),emit_with_api(&ignored,ParserApi::HYBRID_CAPTURES).unwrap());
    for nullable in [false,true] {
        let input=if nullable{format!("// @grameo optimize nullable_failure_heads\n{source}")}else{source.clone()};
        let enabled=grameo::compile_with_options(&input,CompileOptions{optimizations:true,..Default::default()}).unwrap();
        let code=emit_with_api(&enabled,ParserApi::HYBRID_CAPTURES).unwrap();
        assert!(code.contains("fn shared_wrapper_prefix"));assert_eq!(code.matches("fn match_observers_are_noop").count(),1);
        assert!(!code.contains("LazyLock"));
    }
}
