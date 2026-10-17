use grameo_codegen::{ParserApi, generate_with_api};

#[test]
fn generated_offset_entry_uses_the_existing_stack_free_certificate() {
    let grammar = r#"unit={word~";"} word=@{ASCII_ALPHA+}"#;
    for api in [ParserApi::default(), ParserApi::HYBRID_CAPTURES] {
        let generated = generate_with_api(grammar, api).unwrap();
        assert_eq!(generated, generate_with_api(grammar, api).unwrap());
        assert!(generated.contains("pub fn build_captures_at<'g, 'i, B:"));
        assert!(generated.contains("input: &'i str, start_offset: usize, builder: &mut B"));
        assert!(generated.contains("engine::build_stackfree_compact_capture_program_at(self, rule.name(), input, start_offset, ::grameo::MatchOptions::default(), builder)"));
        assert!(!generated.contains("engine::build_compact_capture_program_at("));
        assert!(generated.contains("engine::build_stackfree_compact_capture_program("));
    }
}

#[test]
fn any_grammar_stack_instruction_keeps_undo_even_in_an_unreachable_rule() {
    for expression in [
        "PUSH(\"x\")",
        "PUSH_LITERAL(\"x\")",
        "PEEK[0..]",
        "PEEK",
        "PEEK_ALL",
        "POP",
        "POP_ALL",
        "DROP",
    ] {
        for active in [false, true] {
            let grammar = if active {
                format!("unit={{{expression}}}")
            } else {
                format!("unit={{ANY}} unused={{{expression}}}")
            };
            let generated = generate_with_api(&grammar, ParserApi::HYBRID_CAPTURES).unwrap();
            assert!(generated.contains("pub fn build_captures_at<'g, 'i, B:"));
            assert!(generated.contains("engine::build_compact_capture_program_at(self, rule.name(), input, start_offset, ::grameo::MatchOptions::default(), builder)"), "{grammar}");
            assert!(
                !generated.contains("engine::build_stackfree_compact_capture_program_at("),
                "{grammar}"
            );
            assert!(generated.contains("engine::build_compact_capture_program("));
        }
    }
}

#[test]
fn capture_only_embeddings_do_not_gain_semantic_offset_entry_points() {
    let generated = generate_with_api("unit={ANY}", ParserApi::CAPTURE_ONLY).unwrap();
    assert!(!generated.contains("pub fn build_captures_at<"));
    assert!(!generated.contains("pub fn build_captures<"));
    assert!(!generated.contains("engine::build_stackfree_compact_capture_program_at("));
    assert!(!generated.contains("engine::build_compact_capture_program_at("));
}
