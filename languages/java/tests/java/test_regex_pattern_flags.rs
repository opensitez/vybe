use crate::helpers::run_main;

#[test]
fn pattern_case_insensitive_flag_matches_uppercase_input() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("java", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("JAVA"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_without_case_insensitive_flag_rejects_case_mismatch() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("java"); java.util.regex.Matcher m = p.matcher("JAVA"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_case_insensitive_flag_matches_mixed_case() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("AbC", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_multiline_flag_makes_caret_match_after_newline() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^b", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_without_multiline_flag_caret_only_at_start() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^b"); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_multiline_flag_dollar_matches_before_newline() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a$", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_dotall_flag_dot_matches_newline() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_without_dotall_flag_dot_does_not_match_newline() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b"); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_literal_flag_treats_backslash_d_as_literal() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+", java.util.regex.Pattern.LITERAL); java.util.regex.Matcher m = p.matcher("123"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_literal_flag_matches_exact_literal_text() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b", java.util.regex.Pattern.LITERAL); java.util.regex.Matcher m = p.matcher("a.b"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_comments_flag_ignores_whitespace_in_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a b c", java.util.regex.Pattern.COMMENTS); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_comments_flag_ignores_hash_comment_to_eol() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a # comment\nb", java.util.regex.Pattern.COMMENTS); java.util.regex.Matcher m = p.matcher("ab"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_and_multiline_combined() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^HELLO", java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("x\nhello"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_dotall_and_multiline_combined() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^a.b$", java.util.regex.Pattern.DOTALL | java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("x\na\nb\n"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_and_dotall_combined() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("A.B", java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_flags_method_returns_compiled_flags() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("x", java.util.regex.Pattern.CASE_INSENSITIVE); System.out.println(p.flags() & java.util.regex.Pattern.CASE_INSENSITIVE);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn pattern_flags_zero_for_plain_compile() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("plain"); System.out.println(p.flags());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn pattern_unicode_case_with_case_insensitive_matches_sigma() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("σ", java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE); java.util.regex.Matcher m = p.matcher("Σ"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_canonical_eq_matches_decomposed_and_composed() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\u0065\u0301", java.util.regex.Pattern.CANON_EQ); java.util.regex.Matcher m = p.matcher("\u00e9"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_unicode_character_class_matches_unicode_property() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\p{IsAlphabetic}", java.util.regex.Pattern.UNICODE_CHARACTER_CLASS); java.util.regex.Matcher m = p.matcher("A"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_replace_all() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("cat", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("Cat CAT cAt"); System.out.println(m.replaceAll("dog"));"#,
    );
    assert_eq!(out, vec!["dog dog dog"]);
}

#[test]
fn pattern_multiline_replace_all_on_each_line() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^x", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("x\nx\ny"); System.out.println(m.replaceAll("z"));"#,
    );
    assert_eq!(out, vec!["z", "z", "y"]);
}

#[test]
fn pattern_dotall_find_spans_newline() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("start.*end", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("start\nmiddle\nend"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["start", "middle", "end"]);
}

#[test]
fn pattern_literal_split_on_literal_pipe() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("|", java.util.regex.Pattern.LITERAL); String[] parts = p.split("a|b|c", 3); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn pattern_case_insensitive_looking_at() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("prefix", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("PREFIXextra"); System.out.println(m.lookingAt());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_multiline_count_line_starts() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^line", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("line1\nline2\nline3"); int n = 0; while (m.find()) { n = n + 1; } System.out.println(n);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn pattern_compile_two_arg_overloads_flags() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("abc", java.util.regex.Pattern.CASE_INSENSITIVE); System.out.println(p.pattern());"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn pattern_case_insensitive_group_capture() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(abc)", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("AbC"); m.find(); System.out.println(m.group(1));"#,
    );
    assert_eq!(out, vec!["AbC"]);
}

#[test]
fn pattern_dotall_star_quantifier_crosses_newlines() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.*z", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("a\nb\ncz"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_comments_flag_preserves_escaped_whitespace() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a\\ b", java.util.regex.Pattern.COMMENTS); java.util.regex.Matcher m = p.matcher("a b"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_and_literal_is_literal() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("ABC", java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.LITERAL); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_multiline_does_not_change_dot_behavior_without_dotall() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("a\nb"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_case_insensitive_find_second_match() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("ab", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("Ab x AB"); m.find(); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn pattern_dotall_replace_first_only() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("a\nb a\nb"); System.out.println(m.replaceFirst("X"));"#,
    );
    assert_eq!(out, vec!["X a", "b"]);
}

#[test]
fn pattern_unicode_character_class_digit_property() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\p{Nd}", java.util.regex.Pattern.UNICODE_CHARACTER_CLASS); java.util.regex.Matcher m = p.matcher("7"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_canonical_eq_insensitive_to_order() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\u00e9", java.util.regex.Pattern.CANON_EQ); java.util.regex.Matcher m = p.matcher("\u0065\u0301"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_comments_multiline_pattern_with_comment_lines() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("  foo   # first\n  bar   # second", java.util.regex.Pattern.COMMENTS); java.util.regex.Matcher m = p.matcher("foobar"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_split() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("AND", java.util.regex.Pattern.CASE_INSENSITIVE); String[] parts = p.split("one AND two"); System.out.println(parts.length); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["2", " two"]);
}

#[test]
fn pattern_multiline_start_anchor_on_first_line() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^start", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("start\nend"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_dotall_caret_still_matches_line_start_only() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^a", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("ba"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_literal_caret_is_literal() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^a", java.util.regex.Pattern.LITERAL); java.util.regex.Matcher m = p.matcher("^a"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_matcher_reset() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("ok", java.util.regex.Pattern.CASE_INSENSITIVE); java.util.regex.Matcher m = p.matcher("no"); System.out.println(m.find()); m.reset("OK"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn pattern_multiline_end_anchor_on_last_line() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("end$", java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("start\nend"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_dotall_plus_quantifier() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("x.+y", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("x\n\ny"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_unicode_case_without_case_insensitive_has_no_effect_alone() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("abc", java.util.regex.Pattern.UNICODE_CASE); java.util.regex.Matcher m = p.matcher("ABC"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn pattern_comments_flag_with_escaped_hash() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a\\#b", java.util.regex.Pattern.COMMENTS); java.util.regex.Matcher m = p.matcher("a#b"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_and_multiline_replace_first() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^x", java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.MULTILINE); java.util.regex.Matcher m = p.matcher("x\nX"); System.out.println(m.replaceFirst("y"));"#,
    );
    assert_eq!(out, vec!["y", "X"]);
}

#[test]
fn pattern_literal_dollar_is_literal() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a$", java.util.regex.Pattern.LITERAL); java.util.regex.Matcher m = p.matcher("a$"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_dotall_question_quantifier() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a.b?", java.util.regex.Pattern.DOTALL); java.util.regex.Matcher m = p.matcher("a\n"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_case_insensitive_pattern_constant_value() {
    let out = run_main(r#"System.out.println(java.util.regex.Pattern.CASE_INSENSITIVE == 2);"#);
    assert_eq!(out, vec!["true"]);
}
