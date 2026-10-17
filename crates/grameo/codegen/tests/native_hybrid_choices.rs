use grameo_codegen::{ParserApi, emit_with_api, emit_with_native_hybrid_choices};
#[test]
fn native_choices_require_explicit_generation_and_match_the_gated_fixture() {
    let source = r##"
WHITESPACE = _{ " " }
COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }
entry = { SOI ~ (((a | b) ~ c) | (a ~ b) | c | unsafe_first) ~ EOI }
a = { a_atom ~ PUSH("x") ~ POP }
a_atom = @{ "a" }
b = { b_atom ~ PUSH("z") ~ POP }
b_atom = _{ "b" }
c = { c_alias }
c_alias = !{ c_atom ~ c_tail? }
c_atom = ${ "c" }
c_tail = { "t" }
unsafe_first = { PUSH_LITERAL("q") ~ "q" ~ POP }
"##;
    let grammar = grameo::compile(source).unwrap();
    let ordinary = emit_with_api(&grammar, ParserApi::HYBRID_CAPTURES).unwrap();
    assert!(!ordinary.contains("fn native_hybrid_choices"));
    let native = emit_with_native_hybrid_choices(&grammar, ParserApi::HYBRID_CAPTURES).unwrap();
    assert_eq!(
        native,
        emit_with_native_hybrid_choices(&grammar, ParserApi::HYBRID_CAPTURES).unwrap()
    );
    assert_eq!(
        native,
        include_str!("../../tests/fixtures/native_choices.rs")
    );
}
