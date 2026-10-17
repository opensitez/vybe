const BODY: &str = "root={pad ~ \"x\"}\npad=_{(\"\\n\"|\"\\r\\n\")*}";
#[test]
fn empty_success_default_generation_is_byte_identical() {
    assert_eq!(
        grameo_codegen::generate(BODY).unwrap(),
        grameo_codegen::generate(&format!("// @grameo empty_success_heads pad\n{BODY}")).unwrap()
    );
}
#[test]
fn empty_success_emission_requires_explicit_selection() {
    let g = grameo::compile_with_options(
        &format!("// @grameo empty_success_heads pad\n{BODY}"),
        grameo::CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    let s = grameo_codegen::emit(&g).unwrap();
    assert!(s.contains("fn empty_success_candidate"));
    assert!(s.contains("fn empty_success_metadata_is_current"));
    assert!(
        !grameo_codegen::generate(BODY)
            .unwrap()
            .contains("fn empty_success_candidate")
    );
}
#[test]
fn empty_success_unsafe_shapes_are_rejected() {
    for body in [
        "pad={\"x\"*}",
        "pad=_{\"\"*}",
        "pad=_{other*}\nother={\"x\"}",
        "pad=_{(\"x\"~\"y\")*}",
        "pad=_{\"x\"+}",
        "pad=_{(!\"x\")*}",
    ] {
        assert!(
            grameo::compile_with_options(
                &format!("// @grameo empty_success_heads pad\n{body}"),
                grameo::CompileOptions {
                    optimizations: true,
                    ..Default::default()
                }
            )
            .is_err(),
            "{body}"
        );
    }
}
