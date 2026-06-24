use crate::helpers::run_main;

#[test]
fn char_at_returns_first_code_unit() {
    let out = run_main(r#"String s = "java"; System.out.println(s.charAt(0));"#);
    assert_eq!(out, vec!["j"]);
}

#[test]
fn char_at_reads_middle_code_unit() {
    let out = run_main(r#"String s = "hello"; System.out.println(s.charAt(2));"#);
    assert_eq!(out, vec!["l"]);
}

#[test]
fn char_at_reads_final_code_unit() {
    let out = run_main(r#"String s = "end"; System.out.println(s.charAt(2));"#);
    assert_eq!(out, vec!["d"]);
}

#[test]
fn substring_with_begin_and_end_indices() {
    let out = run_main(r#"String s = "hello"; System.out.println(s.substring(1, 4));"#);
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn substring_zero_width_range_is_empty() {
    let out = run_main(r#"String s = "abc"; System.out.println(s.substring(2, 2));"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn index_of_returns_negative_when_absent() {
    let out = run_main(r#"String s = "foobar"; System.out.println(s.indexOf("z"));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn index_of_with_start_offset_skips_earlier_hits() {
    let out = run_main(r#"String s = "banana"; System.out.println(s.indexOf("na", 3));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn index_of_empty_needle_returns_zero() {
    let out = run_main(r#"String s = "data"; System.out.println(s.indexOf(""));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn index_of_single_character_needle() {
    let out = run_main(r#"String s = "abcde"; System.out.println(s.indexOf("c"));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn last_index_of_finds_rightmost_match() {
    let out = run_main(r#"String s = "banana"; System.out.println(s.lastIndexOf("na"));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn last_index_of_with_from_index_searches_leftward() {
    let out = run_main(r#"String s = "banana"; System.out.println(s.lastIndexOf("a", 3));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn last_index_of_returns_negative_when_absent() {
    let out = run_main(r#"String s = "hello"; System.out.println(s.lastIndexOf("z"));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn starts_with_true_for_matching_prefix() {
    let out = run_main(r#"String s = "prefix"; System.out.println(s.startsWith("pre"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn starts_with_false_for_non_matching_prefix() {
    let out = run_main(r#"String s = "prefix"; System.out.println(s.startsWith("post"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn starts_with_offset_checks_from_index() {
    let out = run_main(r#"String s = "foobar"; System.out.println(s.startsWith("bar", 3));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn ends_with_true_for_matching_suffix() {
    let out = run_main(r#"String s = "filename.txt"; System.out.println(s.endsWith(".txt"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn ends_with_false_for_non_matching_suffix() {
    let out = run_main(r#"String s = "filename.txt"; System.out.println(s.endsWith(".java"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn equals_ignore_case_treats_ascii_case_as_equal() {
    let out = run_main(
        r#"String a = "Java"; String b = "java"; System.out.println(a.equalsIgnoreCase(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn equals_ignore_case_false_when_content_differs() {
    let out = run_main(
        r#"String a = "java"; String b = "kotlin"; System.out.println(a.equalsIgnoreCase(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn compare_to_negative_when_lexicographically_less() {
    let out = run_main(r#"System.out.println("apple".compareTo("banana"));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn compare_to_zero_for_identical_strings() {
    let out = run_main(r#"System.out.println("same".compareTo("same"));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn compare_to_positive_when_lexicographically_greater() {
    let out = run_main(r#"System.out.println("zebra".compareTo("apple"));"#);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn compare_to_ignore_case_ignores_ascii_case() {
    let out = run_main(r#"System.out.println("ABC".compareToIgnoreCase("abc"));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strip_leading_removes_only_leading_whitespace() {
    let out = run_main(r#"String s = "  trim"; System.out.println(s.stripLeading());"#);
    assert_eq!(out, vec!["trim"]);
}

#[test]
fn strip_trailing_removes_only_trailing_whitespace() {
    let out = run_main(r#"String s = "trim  "; System.out.println(s.stripTrailing());"#);
    assert_eq!(out, vec!["trim"]);
}

#[test]
fn strip_removes_both_leading_and_trailing_whitespace() {
    let out = run_main(r#"String s = "  core  "; System.out.println(s.strip());"#);
    assert_eq!(out, vec!["core"]);
}

#[test]
fn replace_swaps_literal_occurrences() {
    let out = run_main(r#"String s = "hello"; System.out.println(s.replace("l", "L"));"#);
    assert_eq!(out, vec!["heLLo"]);
}

#[test]
fn replace_leaves_string_when_no_match() {
    let out = run_main(r#"String s = "abc"; System.out.println(s.replace("z", "Z"));"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn replace_all_substitutes_every_regex_match() {
    let out = run_main(r#"String s = "a1b22c"; System.out.println(s.replaceAll("\\d+", "X"));"#);
    assert_eq!(out, vec!["aXc"]);
}

#[test]
fn replace_first_substitutes_only_first_regex_match() {
    let out = run_main(r##"String s = "a1b2c"; System.out.println(s.replaceFirst("\\d", "#"));"##);
    assert_eq!(out, vec!["a#b2c"]);
}

#[test]
fn split_on_delimiter_returns_segments() {
    let out = run_main(
        r#"String[] parts = "a,b,c".split(","); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn split_with_positive_limit_retains_remainder_in_tail() {
    let out = run_main(
        r#"String[] parts = "a,b,c,d".split(",", 2); System.out.println(parts[0]); System.out.println(parts[1]);"#,
    );
    assert_eq!(out, vec!["a", "b,c,d"]);
}

#[test]
fn split_on_whitespace_regex_groups_words() {
    let out = run_main(
        r#"String[] parts = "one  two\tthree".split("\\s+"); System.out.println(parts[0]); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["one", "three"]);
}

#[test]
fn is_blank_true_for_whitespace_only_string() {
    let out = run_main(r#"String s = "   \t"; System.out.println(s.isBlank());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn is_blank_false_for_visible_characters() {
    let out = run_main(r#"String s = " x "; System.out.println(s.isBlank());"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn is_empty_false_on_nonempty_string() {
    let out = run_main(r#"String s = " "; System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn concat_appends_second_string() {
    let out = run_main(r#"String s = "foo".concat("bar"); System.out.println(s);"#);
    assert_eq!(out, vec!["foobar"]);
}

#[test]
fn repeat_duplicates_string_requested_times() {
    let out = run_main(r#"String s = "ab".repeat(3); System.out.println(s);"#);
    assert_eq!(out, vec!["ababab"]);
}

#[test]
fn repeat_zero_times_yields_empty_string() {
    let out = run_main(r#"String s = "xy".repeat(0); System.out.println(s);"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn value_of_integer_becomes_decimal_digits() {
    let out = run_main(r#"System.out.println(String.valueOf(42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn value_of_boolean_true_becomes_true_text() {
    let out = run_main(r#"System.out.println(String.valueOf(true));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn value_of_char_becomes_single_letter_string() {
    let out = run_main(r#"System.out.println(String.valueOf('Z'));"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn format_inserts_arguments_into_template() {
    let out = run_main(r#"System.out.println(String.format("%s=%d", "count", 7));"#);
    assert_eq!(out, vec!["count=7"]);
}

#[test]
fn formatted_method_fills_placeholder_on_literal() {
    let out = run_main(r#"System.out.println("Hello %s!".formatted("Java"));"#);
    assert_eq!(out, vec!["Hello Java!"]);
}

#[test]
fn matches_returns_true_for_matching_regex() {
    let out = run_main(r#"String s = "abc123"; System.out.println(s.matches(".*\\d+"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn matches_returns_false_for_non_matching_regex() {
    let out = run_main(r#"String s = "letters"; System.out.println(s.matches(".*\\d+"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn value_of_double_formats_decimal_text() {
    let out = run_main(r#"System.out.println(String.valueOf(3.14));"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn to_char_array_exposes_individual_characters() {
    let out = run_main(
        r#"char[] chars = "abc".toCharArray(); System.out.println(chars[0]); System.out.println(chars[2]);"#,
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn hash_code_is_stable_for_same_string_instance() {
    let out = run_main(
        r#"String s = "stable"; int h1 = s.hashCode(); int h2 = s.hashCode(); System.out.println(h1 == h2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn to_string_on_string_returns_same_text() {
    let out = run_main(r#"String s = "literal"; System.out.println(s.toString());"#);
    assert_eq!(out, vec!["literal"]);
}

#[test]
fn char_code_at_returns_utf16_code_unit() {
    let out = run_main(r#"String s = "A"; System.out.println(s.charCodeAt(0));"#);
    assert_eq!(out, vec!["65"]);
}

#[test]
fn code_point_at_returns_bmp_scalar_value() {
    let out = run_main(r#"String s = "B"; System.out.println(s.codePointAt(0));"#);
    assert_eq!(out, vec!["66"]);
}

#[test]
fn join_combines_elements_with_delimiter() {
    let out = run_main(r#"System.out.println(String.join("-", "a", "b", "c"));"#);
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn contains_empty_substring_is_always_true() {
    let out = run_main(r#"String s = "any"; System.out.println(s.contains(""));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn length_counts_utf16_code_units_not_graphemes() {
    let out = run_main(r#"String s = "caf\u00e9"; System.out.println(s.length());"#);
    assert_eq!(out, vec!["4"]);
}
