use crate::helpers::run_main;

#[test]
fn stringtokenizer_next_token_reads_first_word() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("alpha beta"); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["alpha"]);
}

#[test]
fn stringtokenizer_next_token_reads_second_after_first() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("alpha beta"); st.nextToken(); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["beta"]);
}

#[test]
fn stringtokenizer_next_token_reads_third_in_sequence() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("one two three"); st.nextToken(); st.nextToken(); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["three"]);
}

#[test]
fn stringtokenizer_has_more_tokens_true_at_start() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a b"); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringtokenizer_has_more_tokens_false_after_last() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("solo"); st.nextToken(); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stringtokenizer_has_more_tokens_true_with_remaining() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("x y"); st.nextToken(); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringtokenizer_count_tokens_on_three_word_input() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a b c"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_count_tokens_decrements_after_next() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a b c"); st.nextToken(); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringtokenizer_count_tokens_single_token() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("only"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringtokenizer_has_more_elements_via_enumeration() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("p q"); System.out.println(st.hasMoreElements());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringtokenizer_next_element_returns_first_token() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("first second"); System.out.println(st.nextElement());"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn stringtokenizer_next_element_advances_like_next_token() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("first second"); st.nextElement(); System.out.println(st.nextElement());"#,
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn stringtokenizer_comma_delimiter_splits_csv() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("red,green,blue", ","); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["red", "green"]);
}

#[test]
fn stringtokenizer_comma_delimiter_reads_last_field() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("x,y,z", ","); st.nextToken(); st.nextToken(); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn stringtokenizer_comma_count_tokens() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b,c,d", ","); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringtokenizer_tab_delimiter_splits_fields() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("col1\tcol2\tcol3", "\t"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["col1", "col2"]);
}

#[test]
fn stringtokenizer_semicolon_delimiter_splits_record() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("10;20;30", ";"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn stringtokenizer_pipe_delimiter_splits_segments() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("x|y|z", "|"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn stringtokenizer_colon_delimiter_splits_key_value() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("host:8080", ":"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["host", "8080"]);
}

#[test]
fn stringtokenizer_slash_delimiter_splits_path() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("usr/local/bin", "/"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["usr", "local"]);
}

#[test]
fn stringtokenizer_hyphen_delimiter_splits_dashed_text() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("vybe-java-test", "-"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["vybe", "java"]);
}

#[test]
fn stringtokenizer_at_sign_delimiter_splits_email_like() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("user@host", "@"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["user", "host"]);
}

#[test]
fn stringtokenizer_mixed_delimiter_chars_in_delim_set() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b;c|d", ",;|"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn stringtokenizer_mixed_delimiters_count_all_tokens() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b;c|d", ",;|"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringtokenizer_empty_input_has_no_tokens() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer(""); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stringtokenizer_empty_input_count_is_zero() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer(""); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringtokenizer_leading_delimiter_skips_to_first_token() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer(",a,b", ","); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn stringtokenizer_trailing_delimiter_yields_no_extra_token() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b,", ","); st.nextToken(); st.nextToken(); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stringtokenizer_leading_whitespace_skipped_by_default() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("   hello world"); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn stringtokenizer_consecutive_commas_skip_empty_fields() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,,b", ","); st.nextToken(); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn stringtokenizer_consecutive_spaces_treated_as_one() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a   b   c"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_single_token_no_delimiter() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("standalone"); System.out.println(st.nextToken()); System.out.println(st.hasMoreTokens());"#,
    );
    assert_eq!(out, vec!["standalone", "false"]);
}

#[test]
fn stringtokenizer_while_loop_concatenates_three_tokens() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("do re mi"); String s = ""; while (st.hasMoreTokens()) { s = s + st.nextToken(); } System.out.println(s);"#,
    );
    assert_eq!(out, vec!["doremi"]);
}

#[test]
fn stringtokenizer_while_loop_counts_comma_separated() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("1,2,3,4,5", ","); int n = 0; while (st.hasMoreTokens()) { st.nextToken(); n = n + 1; } System.out.println(n);"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stringtokenizer_return_delimiters_false_omits_comma() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b", ",", false); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn stringtokenizer_return_delimiters_true_includes_comma() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b", ",", true); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["a", ","]);
}

#[test]
fn stringtokenizer_return_delimiters_true_preserves_space() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a b", " ", true); st.nextToken(); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec![" "]);
}

#[test]
fn stringtokenizer_return_delimiters_true_count_includes_separators() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a,b", ",", true); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_two_arg_constructor_custom_delim() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("x-y-z", "-"); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn stringtokenizer_three_arg_constructor_with_return_flag() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a|b", "|", true); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["a", "|"]);
}

#[test]
fn stringtokenizer_default_delim_splits_tab_and_space() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a\tb"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringtokenizer_default_delim_splits_newline() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("line1\nline2"); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["line1", "line2"]);
}

#[test]
fn stringtokenizer_multiple_chars_in_delimiter_set() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("a b,c;d", " ,;"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringtokenizer_next_token_after_count_depleted() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("one", " "); System.out.println(st.nextToken()); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["one", "0"]);
}

#[test]
fn stringtokenizer_has_more_elements_false_at_end() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("z"); st.nextElement(); System.out.println(st.hasMoreElements());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stringtokenizer_pipe_separated_path_token_count() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("src|main|java", "|"); int n = 0; while (st.hasMoreTokens()) { st.nextToken(); n++; } System.out.println(n);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_colon_separated_time_parts() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("12:30:45", ":"); System.out.println(st.nextToken()); System.out.println(st.nextToken()); System.out.println(st.nextToken());"#,
    );
    assert_eq!(out, vec!["12", "30", "45"]);
}

#[test]
fn stringtokenizer_hyphenated_locale_tag_parts() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("en-US-UTF8", "-"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_at_delimited_triple_split() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("name@domain@tld", "@"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringtokenizer_slash_path_skips_empty_leading() {
    let out = run_main(
        r#"java.util.StringTokenizer st = new java.util.StringTokenizer("/a/b/c/", "/"); System.out.println(st.countTokens());"#,
    );
    assert_eq!(out, vec!["3"]);
}
