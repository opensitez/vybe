use crate::helpers::run_main;

#[test]
fn split_on_comma_returns_three_parts() {
    let out = run_main(
        r#"String[] parts = "a,b,c".split(","); System.out.println(parts.length); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", "a", "c"]);
}

#[test]
fn split_on_regex_digit_run_separates_letters() {
    let out = run_main(
        r#"String[] parts = "a1b22c".split("\\d+"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn split_on_whitespace_regex_groups_words() {
    let out = run_main(
        r#"String[] parts = "one  two\tthree".split("\\s+"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["one", "three"]);
}

#[test]
fn split_on_dot_regex_splits_extension() {
    let out = run_main(
        r#"String[] parts = "archive.tar.gz".split("\\."); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["archive", "gz"]);
}

#[test]
fn split_on_pipe_regex_splits_fields() {
    let out = run_main(
        r#"String[] parts = "x|y|z".split("\\|"); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["y"]);
}

#[test]
fn split_on_colon_regex_parses_time_parts() {
    let out = run_main(
        r#"String[] parts = "12:30:45".split(":"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["12", "45"]);
}

#[test]
fn split_on_character_class_regex_splits_vowels() {
    let out = run_main(
        r#"String[] parts = "brisk".split("[aeiou]"); System.out.println(parts.length);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn split_with_positive_limit_two_retains_remainder_in_tail() {
    let out = run_main(
        r#"String[] parts = "a,b,c,d".split(",", 2); System.out.println(parts.length); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["2", "a", "b,c,d"]);
}

#[test]
fn split_with_limit_three_splits_first_two_only() {
    let out = run_main(
        r#"String[] parts = "a,b,c,d".split(",", 3); System.out.println(parts.length); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", "c,d"]);
}

#[test]
fn split_with_limit_one_returns_whole_input() {
    let out = run_main(
        r#"String[] parts = "a,b,c".split(",", 1); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "a,b,c"]);
}

#[test]
fn split_with_limit_zero_drops_trailing_empty_tokens() {
    let out = run_main(
        r#"String[] parts = "a,b,".split(",", 0); System.out.println(parts.length);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn split_with_negative_limit_keeps_trailing_empty_token() {
    let out = run_main(
        r#"String[] parts = "a,b,".split(",", -1); System.out.println(parts.length); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", ""]);
}

#[test]
fn split_preserve_trailing_empty_with_negative_limit_on_double_comma() {
    let out = run_main(
        r#"String[] parts = "a,,b".split(",", -1); System.out.println(parts.length); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["3", ""]);
}

#[test]
fn split_preserve_multiple_trailing_empties_with_negative_limit() {
    let out = run_main(
        r#"String[] parts = "a,b,,".split(",", -1); System.out.println(parts.length); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["4", ""]);
}

#[test]
fn split_trailing_empty_dropped_with_default_limit() {
    let out = run_main(
        r#"String[] parts = "x,y,".split(","); System.out.println(parts.length);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn split_trailing_empty_preserved_with_negative_limit_on_regex() {
    let out = run_main(
        r#"String[] parts = "a b ".split("\\s+", -1); System.out.println(parts.length); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", ""]);
}

#[test]
fn split_on_regex_with_limit_two_keeps_tail() {
    let out = run_main(
        r#"String[] parts = "a1b2c3".split("\\d", 2); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["a", "b2c3"]);
}

#[test]
fn split_single_token_without_delimiter_yields_one_element() {
    let out = run_main(
        r#"String[] parts = "solo".split(","); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "solo"]);
}

#[test]
fn split_empty_string_on_comma_yields_single_empty_element() {
    let out = run_main(
        r#"String[] parts = "".split(","); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", ""]);
}

#[test]
fn split_on_regex_plus_quantifier_collapses_runs() {
    let out = run_main(
        r#"String[] parts = "a--b__c".split("[-_]+"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn split_on_alternation_regex_picks_any_delimiter() {
    let out = run_main(
        r#"String[] parts = "a,b;c|d".split("[,;|]"); System.out.println(parts.length); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["4", "d"]);
}

#[test]
fn split_with_limit_zero_on_trailing_delimiter_drops_empty() {
    let out = run_main(
        r#"String[] parts = "one,two,".split(",", 0); System.out.println(parts.length); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["2", "two"]);
}

#[test]
fn split_with_negative_limit_on_leading_delimiter_preserves_leading_empty() {
    let out = run_main(
        r#"String[] parts = ",a,b".split(",", -1); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["", "a"]);
}

#[test]
fn split_on_word_boundary_regex_between_tokens() {
    let out = run_main(
        r#"String[] parts = "fooBar".split("(?=[A-Z])"); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["foo", "Bar"]);
}

#[test]
fn split_on_equals_sign_in_key_value_pairs() {
    let out = run_main(
        r#"String[] parts = "k1=v1&k2=v2".split("[=&]"); System.out.println(parts[0]); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["k1", "v2"]);
}

#[test]
fn split_limit_four_on_five_field_csv() {
    let out = run_main(
        r#"String[] parts = "a,b,c,d,e".split(",", 4); System.out.println(parts.length); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["4", "d,e"]);
}

#[test]
fn split_regex_digit_between_letters_with_negative_limit() {
    let out = run_main(
        r#"String[] parts = "a1".split("\\d", -1); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["a", ""]);
}

#[test]
fn split_on_slash_regex_for_path_segments() {
    let out = run_main(
        r#"String[] parts = "/usr/local/bin".split("/"); System.out.println(parts[1]); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["usr", "bin"]);
}

#[test]
fn split_preserve_trailing_empty_on_repeated_delimiter() {
    let out = run_main(
        r#"String[] parts = "a,,b,".split(",", -1); System.out.println(parts.length); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["4", ""]);
}

#[test]
fn split_on_tab_regex_splits_columns() {
    let out = run_main(
        r#"String[] parts = "name\tage\tcity".split("\\t"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["name", "city"]);
}

#[test]
fn split_with_limit_two_on_regex_whitespace() {
    let out = run_main(
        r#"String[] parts = "aa bb cc".split("\\s+", 2); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["aa", "bb cc"]);
}

#[test]
fn split_on_non_digit_regex_extracts_numbers() {
    let out = run_main(
        r#"String[] parts = "12a34b56".split("\\D+"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["12", "56"]);
}

#[test]
fn split_default_limit_drops_only_trailing_empty_not_internal() {
    let out = run_main(
        r#"String[] parts = "a,,b".split(","); System.out.println(parts.length); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["3", ""]);
}

#[test]
fn split_negative_limit_keeps_internal_and_trailing_empty() {
    let out = run_main(
        r#"String[] parts = "a,,b,".split(",", -1); System.out.println(parts.length); System.out.println(parts[3]);"#,
    );
    assert_eq!(out, vec!["4", ""]);
}

#[test]
fn split_on_regex_anchors_start_of_line() {
    let out = run_main(
        r#"String[] parts = "id:42".split("^id:"); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["", "42"]);
}

#[test]
fn split_on_comma_with_spaces_regex_trims_gaps() {
    let out = run_main(
        r#"String[] parts = "a, b , c".split(",\\s*"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn split_limit_zero_equals_default_for_trailing_empty() {
    let out = run_main(
        r#"String[] a = "x,y,".split(",", 0); String[] b = "x,y,".split(","); System.out.println(a.length == b.length);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn split_on_regex_question_mark_optional_delimiter() {
    let out = run_main(
        r#"String[] parts = "a-b_c".split("[-_]?"); System.out.println(parts.length);"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn split_preserve_trailing_empty_on_regex_delimiter_only_at_end() {
    let out = run_main(
        r#"String[] parts = "alpha-".split("-", -1); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["alpha", ""]);
}

#[test]
fn split_on_regex_with_limit_one_keeps_entire_string() {
    let out = run_main(
        r#"String[] parts = "a:b:c".split(":", 1); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "a:b:c"]);
}
