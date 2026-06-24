use crate::helpers::run_main;

#[test]
fn pattern_compile_accepts_digit_regex() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); System.out.println(p != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn pattern_matcher_returns_non_null_for_input() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_find_locates_first_digit_run() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("ab12cd"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_find_returns_false_when_pattern_absent() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("letters"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn matcher_find_locates_second_match_on_second_call() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d"); java.util.regex.Matcher m = p.matcher("a1b2"); m.find(); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_matches_true_for_entire_string_match() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[a-z]+"); java.util.regex.Matcher m = p.matcher("hello"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_matches_false_when_extra_characters_remain() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[a-z]+"); java.util.regex.Matcher m = p.matcher("hello1"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn matcher_looking_at_true_for_prefix_match() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("pre"); java.util.regex.Matcher m = p.matcher("prefix"); System.out.println(m.lookingAt());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_looking_at_false_when_pattern_not_at_start() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("fix"); java.util.regex.Matcher m = p.matcher("suffix"); System.out.println(m.lookingAt());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn matcher_group_zero_returns_full_match_text() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("x99y"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn matcher_group_one_returns_first_capturing_parentheses() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\d+)-(\\d+)"); java.util.regex.Matcher m = p.matcher("12-34"); m.find(); System.out.println(m.group(1));"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn matcher_group_two_returns_second_capturing_parentheses() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\d+)-(\\d+)"); java.util.regex.Matcher m = p.matcher("12-34"); m.find(); System.out.println(m.group(2));"#,
    );
    assert_eq!(out, vec!["34"]);
}

#[test]
fn matcher_replace_all_substitutes_every_digit_run() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("a1b22c"); System.out.println(m.replaceAll("#"));"##,
    );
    assert_eq!(out, vec!["a#b#c"]);
}

#[test]
fn matcher_replace_all_with_backreference_reorders_groups() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\w+)-(\\w+)"); java.util.regex.Matcher m = p.matcher("foo-bar"); System.out.println(m.replaceAll("$2_$1"));"#,
    );
    assert_eq!(out, vec!["bar_foo"]);
}

#[test]
fn pattern_split_on_comma_returns_three_segments() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("a,b,c"); System.out.println(parts.length); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["3", "b"]);
}

#[test]
fn pattern_split_with_limit_two_retains_tail() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("a,b,c,d", 2); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["a", "b,c,d"]);
}

#[test]
fn pattern_split_on_whitespace_groups_words() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\s+"); String[] parts = p.split("one  two\tthree"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["one", "three"]);
}

#[test]
fn matcher_find_on_email_like_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[a-z]+@[a-z]+"); java.util.regex.Matcher m = p.matcher("user@host"); System.out.println(m.find());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_group_zero_on_word_boundary_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\bjava\\b"); java.util.regex.Matcher m = p.matcher("run java now"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn matcher_matches_on_exact_ip_octet_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d{1,3}"); java.util.regex.Matcher m = p.matcher("192"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_looking_at_on_optional_sign_number() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("-?\\d+"); java.util.regex.Matcher m = p.matcher("-42px"); System.out.println(m.lookingAt());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_find_counts_three_digit_tokens() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("1a22b333"); int n = 0; while (m.find()) { n = n + 1; } System.out.println(n);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn matcher_replace_all_strips_leading_zeros() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^0+"); java.util.regex.Matcher m = p.matcher("007"); System.out.println(m.replaceAll(""));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn pattern_split_preserves_trailing_empty_with_negative_limit() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("a,b,", -1); System.out.println(parts.length); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", ""]);
}

#[test]
fn pattern_split_with_zero_limit_drops_trailing_empty() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("a,b,", 0); System.out.println(parts.length);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn matcher_group_on_alternation_first_branch() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(cat|dog)"); java.util.regex.Matcher m = p.matcher("dog"); m.find(); System.out.println(m.group(1));"#,
    );
    assert_eq!(out, vec!["dog"]);
}

#[test]
fn matcher_matches_false_for_empty_input_with_plus_quantifier() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(".+"); java.util.regex.Matcher m = p.matcher(""); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn matcher_find_on_dot_star_finds_empty_match_at_start() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(".*"); java.util.regex.Matcher m = p.matcher("abc"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn pattern_compile_literal_dot_matches_dot_only() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\."); java.util.regex.Matcher m = p.matcher("a.b"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["."]);
}

#[test]
fn matcher_replace_all_normalizes_multiple_spaces() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(" +"); java.util.regex.Matcher m = p.matcher("a  b   c"); System.out.println(m.replaceAll(" "));"#,
    );
    assert_eq!(out, vec!["a b c"]);
}

#[test]
fn pattern_split_on_pipe_delimiter() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\|"); String[] parts = p.split("x|y|z"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["x", "z"]);
}

#[test]
fn matcher_looking_at_false_for_case_sensitive_mismatch() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("Java"); java.util.regex.Matcher m = p.matcher("java"); System.out.println(m.lookingAt());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn matcher_find_on_hex_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[0-9a-f]+"); java.util.regex.Matcher m = p.matcher("color:ff00aa"); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn matcher_replace_all_wraps_words_with_brackets() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\w+"); java.util.regex.Matcher m = p.matcher("hi there"); System.out.println(m.replaceAll("[$0]"));"#,
    );
    assert_eq!(out, vec!["[hi] [there]"]);
}

#[test]
fn pattern_split_single_element_without_delimiter() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("solo"); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "solo"]);
}

#[test]
fn matcher_matches_on_true_literal_pattern() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("true"); java.util.regex.Matcher m = p.matcher("true"); System.out.println(m.matches());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matcher_find_after_partial_scan_continues_from_end() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("aba"); m.find(); m.find(); System.out.println(m.group(0));"#,
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn pattern_split_limit_one_returns_whole_input() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile(","); String[] parts = p.split("a,b,c", 1); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "a,b,c"]);
}

#[test]
fn matcher_group_on_nested_quantifier_match() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(a+)b"); java.util.regex.Matcher m = p.matcher("aaab"); m.find(); System.out.println(m.group(1));"#,
    );
    assert_eq!(out, vec!["aaa"]);
}

#[test]
fn matcher_replace_all_removes_non_alphanumeric_runs() {
    let out = run_main(
        r#"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[^a-z]+"); java.util.regex.Matcher m = p.matcher("a--b__c"); System.out.println(m.replaceAll(""));"#,
    );
    assert_eq!(out, vec!["abc"]);
}
