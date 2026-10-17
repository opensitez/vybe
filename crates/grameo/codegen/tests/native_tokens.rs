use grameo::{CompileOptions, optimization::OptimizationSelection};
fn opts() -> CompileOptions {
    CompileOptions {
        optimizations: true,
        ..Default::default()
    }
}
#[test]
fn native_tokens_preserve_default_selection_and_public_selection_literal() {
    let _old = OptimizationSelection {
        ascii_xid: false,
        nullable_failure_heads: false,
        native_choices: false,
        native_rule_heads: false,
    };
    let body = "token=@{ASCII_ALPHA ~ (ASCII_ALPHANUMERIC | \"_\")*}";
    let annotated = format!("// @grameo optimize native_tokens\n{body}");
    assert!(
        !grameo::compile(&annotated)
            .unwrap()
            .native_token_compilation()
    );
    assert!(
        grameo::compile_with_options(&annotated, opts())
            .unwrap()
            .native_token_compilation()
    );
    assert_eq!(
        grameo_codegen::generate(body).unwrap(),
        grameo_codegen::generate(&annotated).unwrap()
    );
}
#[test]
fn native_tokens_conservatively_certify_regular_ascii_and_utf8_any_and_decline_unknown_effects() {
    for (s, want) in [
        ("token=@{ASCII_DIGIT+}", true),
        (
            "token=@{\"0x\" ~ ASCII_HEX_DIGIT+ | ASCII_DIGIT+ ~ (\".\" ~ ASCII_DIGIT*)?}",
            true,
        ),
        (
            "token=@{!\"for\" ~ ASCII_ALPHA ~ (ASCII_ALPHANUMERIC | \"_\")*}",
            true,
        ),
        ("token=@{x}\nx=_{\"x\"}", true),
        ("token=@{PUSH(\"x\") ~ POP}", false),
        ("token=@{x}\nx=!{\"x\"}", false),
        ("token=@{x}\nx=${\"x\"}", false),
        ("token=@{\"λ\"}", false),
        ("token=@{ANY+}", true),
        ("token=@{(\"\"){0,3}}", false),
        ("token=@{\"(\" ~ token? ~ \")\" | \"x\"}", false),
        ("token=${\"x\"}", false),
    ] {
        let grammar = grameo::compile_with_options(
            &format!("// @grameo optimize native_tokens\n{s}"),
            opts(),
        )
        .unwrap();
        let a = grameo_codegen::emit(&grammar).unwrap();
        assert_eq!(a.contains("fn native_token("), want, "{s}");
        assert_eq!(
            a.contains("fn native_token_metadata_is_current"),
            want,
            "{s}"
        );
        assert_eq!(a, grameo_codegen::emit(&grammar).unwrap());
    }
}
#[test]
fn native_tokens_compose_with_other_specializations_without_duplicate_observer_methods() {
    for extras in [
        "",
        "nullable_failure_heads",
        "native_choices native_rule_heads",
        "nullable_failure_heads native_choices native_rule_heads",
    ] {
        let source = format!(
            "// @grameo optimize native_tokens {extras}\nentry={{a|b}} a=@{{ASCII_ALPHA+}} b=@{{ASCII_DIGIT+}}"
        );
        let g = grameo::compile_with_options(&source, opts()).unwrap();
        let s = grameo_codegen::emit(&g).unwrap();
        assert!(s.contains("fn native_token("));
        assert_eq!(
            s.matches("fn match_observers_are_noop").count(),
            1,
            "{extras}"
        );
    }
}
#[test]
fn selected_choice_heads_certify_positive_lookahead_without_claiming_negative_or_empty_heads() {
    use grameo::program::Program;
    for (head, known) in [("&\"$\"", true), ("!\"$\"", false), ("&\"\"", false)] {
        let body = format!(
            "select={{{head} ~ variable | number}} variable=@{{\"$\" ~ ASCII_ALPHA+}} number=@{{ASCII_DIGIT+}}"
        );
        let source = format!("// @grameo choice_heads select\n{body}");
        let g = grameo::compile_with_options(&source, opts()).unwrap();
        let id = g.rule(g.rule_id("select").unwrap()).expression;
        let s = grameo_codegen::emit(&g).unwrap();
        assert_eq!(s.contains(&format!("({id},0)=>Some")), known, "{head}");
        assert!(s.contains("fn choice_head_branch_metadata_is_current"));
    }
}

#[test]
fn generic_pratt_fixtures_are_current_codegen_output() {
    let body = include_str!("../../tests/fixtures/native_tokens.grammar");
    let reference = grameo::compile_with_options(body, opts()).unwrap();
    let candidate = grameo::compile_with_options(
        &format!("// @grameo optimize native_tokens\n// @grameo choice_heads atom\n{body}"),
        opts(),
    )
    .unwrap();
    assert_eq!(
        grameo_codegen::emit_with_api(&reference, grameo_codegen::ParserApi::HYBRID_CAPTURES)
            .unwrap(),
        include_str!("../../tests/fixtures/native_tokens_reference.rs")
    );
    assert_eq!(
        grameo_codegen::emit_with_api(&candidate, grameo_codegen::ParserApi::HYBRID_CAPTURES)
            .unwrap(),
        include_str!("../../tests/fixtures/native_tokens_candidate.rs")
    );
}

#[test]
fn utf8_any_certifies_comment_prefixes_without_enabling_nonlexical_effects(){
 let s=r#"// @grameo optimize native_tokens
var=@{"$"~ASCII_ALPHA+} space=_{" "|"\t"} comment=_{"/*"~(!"*/"~ANY)*~"*/"} prefix=@{var~(space|comment)*~"["}"#;
 let g=grameo::compile_with_options(s,opts()).unwrap();let out=grameo_codegen::emit(&g).unwrap();
 use grameo::program::Program;let prefix=g.rule_id("prefix").unwrap();
 assert!(out.contains(&format!("{prefix} if remaining_depth>=")));
 assert!(out.contains("std::str::from_utf8(s.get(p..end)?).ok()?"));
 assert!(out.contains(" && source.is_char_boundary(offset)=>"));
 assert!(!out.contains("if !source.is_char_boundary(offset){return None;}"));
 assert_eq!(grameo_codegen::generate(s).unwrap(),grameo_codegen::generate(s.lines().skip(1).collect::<Vec<_>>().join("\n").as_str()).unwrap());
}

use grameo::program::Program as _;

#[path = "../../tests/fixtures/native_token_failures_candidate.rs"]
#[allow(dead_code, unused_comparisons)]
mod negative_tokens;

#[test]
fn negative_tokens_are_explicit_and_ignored_by_default() {
    let body = include_str!("../../tests/fixtures/native_token_failures.grammar");
    let source = format!("// @grameo optimize native_token_failures\n{body}");
    assert!(!grameo::compile(&source).unwrap().native_token_failures());
    let selected = grameo::compile_with_options(&source, opts()).unwrap();
    assert!(selected.native_token_compilation());
    assert!(selected.native_token_failures());
    assert_eq!(grameo_codegen::generate(body).unwrap(), grameo_codegen::generate(&source).unwrap());
    assert_eq!(grameo_codegen::emit_with_api(&selected, grameo_codegen::ParserApi::HYBRID_CAPTURES).unwrap(), include_str!("../../tests/fixtures/native_token_failures_candidate.rs"));
}

#[test]
fn known_failure_is_distinct_from_unknown_depth_and_invalid_boundary() {
    use grameo::program::Program;
    let p = negative_tokens::Parser;
    let scalar = p.rule_id("scalar").unwrap();
    let ascii = p.rule_id("ascii").unwrap();
    let not_x = p.rule_id("not_x").unwrap();
    assert_eq!(p.native_token_result(scalar, "", 0, usize::MAX), Some(None));
    assert_eq!(p.native_token_result(scalar, "😀", 0, usize::MAX), Some(Some(4)));
    assert_eq!(p.native_token_result(ascii, "7", 0, usize::MAX), Some(None));
    assert_eq!(p.native_token_result(usize::MAX, "x", 0, usize::MAX), None);
    assert_eq!(p.native_token_result(scalar, "x", 0, 0), None);
    for rule in [scalar, ascii, not_x] {
        assert_eq!(p.native_token_result(rule, "é", 1, usize::MAX), None);
        assert_eq!(p.native_token_result(rule, "é", usize::MAX, usize::MAX), None);
    }
}

#[derive(Debug)]
struct LegacyPositive;
impl grameo::program::Program for LegacyPositive {
    fn rule_id(&self, name: &str) -> Option<usize> { negative_tokens::Parser.rule_id(name) }
    fn rule(&self, rule: usize) -> grameo::program::RuleSpec<'_> { negative_tokens::Parser.rule(rule) }
    fn instruction(&self, id: usize) -> grameo::program::Instruction<'_> { negative_tokens::Parser.instruction(id) }
    fn native_token(&self, rule: usize, source: &str, offset: usize, depth: usize) -> Option<usize> { negative_tokens::Parser.native_token(rule, source, offset, depth) }
    fn native_token_metadata_is_current(&self) -> bool { true }
    fn atomic_expression(&self, id: usize) -> bool { negative_tokens::Parser.atomic_expression(id) }
    fn match_observers_are_noop(&self) -> bool { true }
}
#[test]
fn legacy_positive_provider_never_claims_a_negative_result() {
    use grameo::program::Program;
    let p = LegacyPositive;
    let scalar = p.rule_id("scalar").unwrap();
    assert_eq!(p.native_token_result(scalar, "a", 0, usize::MAX), Some(Some(1)));
    assert_eq!(p.native_token_result(scalar, "", 0, usize::MAX), None);
}

#[test]
fn ascii_positive_roots_keep_no_any_boundary_guard() {
    use grameo::program::Program;
    let g = grameo::compile_with_options("// @grameo optimize native_tokens\nascii=@{ASCII_ALPHA+} scalar=@{ANY} negative=@{!ANY}", opts()).unwrap();
    let out = grameo_codegen::emit(&g).unwrap();
    let arm = |name: &str| { let id = g.rule_id(name).unwrap(); out.lines().find(|line| line.trim_start().starts_with(&format!("{id} if remaining_depth>="))).unwrap() };
    assert!(!arm("ascii").contains("is_char_boundary"));
    assert!(arm("scalar").contains("is_char_boundary"));
    assert!(arm("negative").contains("is_char_boundary"));
    assert!(!out.contains("if !source.is_char_boundary(offset){return None;}"));
}
#[test]
fn exact_quote_and_delimited_proofs_retain_specialized_priority() {
    use grameo::program::Program;
    let p = negative_tokens::Parser;
    for name in ["quoted", "comment"] {
        let rule = p.rule_id(name).unwrap();
        assert!(p.quoted_fragment(rule).is_some() || p.delimited_fragment(rule).is_some());
        assert_eq!(p.native_token_result(rule, "'é'", 0, usize::MAX), None);
    }
    assert_eq!(p.parse(negative_tokens::Rule::quoted, "'é😀'").unwrap().consumed(), 8);
    assert_eq!(p.parse(negative_tokens::Rule::comment, "/*é😀*/").unwrap().consumed(), 10);
}

#[test]
fn negative_failure_preserves_captures_error_precedence_and_finite_limits() {
    fn shape(tree: grameo::ParseTree<'_, '_>) -> (usize, Vec<(String, grameo::Span, usize)>) {
        let consumed = tree.consumed();
        let mut captures = Vec::new();
        let mut pending = tree.pairs().map(|p| (p, 0)).collect::<Vec<_>>();
        while let Some((p, depth)) = pending.pop() {
            captures.push((p.rule_name().to_owned(), p.as_span(), depth));
            pending.extend(p.into_inner().map(|p| (p, depth + 1)));
        }
        (consumed, captures)
    }
    for source in ["$ordinary", "$a[1]", "$a[", "$a[é]"] {
        for max_steps in [0, 1, 64, usize::MAX] {
            for max_rule_depth in [0, 1, 4, usize::MAX] {
                let options = grameo::MatchOptions { max_steps, max_rule_depth };
                let positive = grameo::engine::parse_program(&LegacyPositive, "entry", source, options).map(shape);
                let negative = grameo::engine::parse_program(&negative_tokens::Parser, "entry", source, options).map(shape);
                assert_eq!(positive, negative, "{source:?}/{max_steps}/{max_rule_depth}");
            }
        }
    }
}
