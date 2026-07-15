use crate::helpers::{run_in_main, run_main};

#[test]
fn region_matches_true_for_identical_prefix_at_zero() {
    let out = run_main(
        r#"String a = "hello world"; String b = "hello"; System.out.println(a.regionMatches(0, b, 0, 5));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn region_matches_false_when_other_region_too_short() {
    let out = run_main(
        r#"String a = "abc"; String b = "abcdef"; System.out.println(a.regionMatches(0, b, 0, 5));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn region_matches_true_with_nonzero_toffset() {
    let out = run_main(
        r#"String a = "prefix:data"; String b = "data"; System.out.println(a.regionMatches(7, b, 0, 4));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn region_matches_false_on_case_sensitive_mismatch() {
    let out = run_main(
        r#"String a = "Java"; String b = "java"; System.out.println(a.regionMatches(0, b, 0, 4));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn region_matches_ignore_case_true_for_ascii_difference() {
    let out = run_main(
        r#"String a = "Java"; String b = "java"; System.out.println(a.regionMatches(true, 0, b, 0, 4));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn region_matches_ignore_case_false_when_letters_differ() {
    let out = run_main(
        r#"String a = "java"; String b = "jaba"; System.out.println(a.regionMatches(true, 0, b, 0, 4));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn region_matches_with_both_offsets_compares_interior_slices() {
    let out = run_main(
        r#"String a = "xxabcyy"; String b = "zzabcww"; System.out.println(a.regionMatches(2, b, 2, 3));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_by_code_points_advances_by_positive_count() {
    let out = run_main(r#"String s = "abcd"; System.out.println(s.offsetByCodePoints(1, 2));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn offset_by_code_points_retreats_with_negative_count() {
    let out = run_main(r#"String s = "abcd"; System.out.println(s.offsetByCodePoints(3, -2));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_by_code_points_counts_supplementary_as_single_step() {
    let out =
        run_main(r#"String s = "a\uD83D\uDE00b"; System.out.println(s.offsetByCodePoints(0, 2));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn code_point_at_reads_ascii_letter_scalar() {
    let out = run_main(r#"String s = "Z"; System.out.println(s.codePointAt(0));"#);
    assert_eq!(out, vec!["90"]);
}

#[test]
fn code_point_at_reads_emoji_supplementary_scalar() {
    let out = run_main(r#"String s = "\uD83D\uDE00"; System.out.println(s.codePointAt(0));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn code_point_at_after_supplementary_uses_trailing_unit_index() {
    let out = run_main(r#"String s = "x\uD83D\uDE00y"; System.out.println(s.codePointAt(2));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn code_point_before_reads_prior_scalar_value() {
    let out = run_main(r#"String s = "ab"; System.out.println(s.codePointBefore(2));"#);
    assert_eq!(out, vec!["98"]);
}

#[test]
fn code_point_count_equals_length_for_bmp_only_string() {
    let out = run_main(r#"String s = "four"; System.out.println(s.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn code_point_count_treats_supplementary_pair_as_one() {
    let out =
        run_main(r#"String s = "a\uD83D\uDE00b"; System.out.println(s.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn code_point_count_on_subrange_excludes_outside_slice() {
    let out = run_main(r#"String s = "abcdef"; System.out.println(s.codePointCount(1, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn get_bytes_default_charset_matches_ascii_length() {
    let out = run_main(r#"byte[] data = "abc".getBytes(); System.out.println(data.length);"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn get_bytes_utf8_charset_preserves_ascii_bytes() {
    let out = run_main(
        r#"byte[] data = "hi".getBytes(java.nio.charset.StandardCharsets.UTF_8); System.out.println(data[0]); System.out.println(data[1]);"#,
    );
    assert_eq!(out, vec!["104", "105"]);
}

#[test]
fn replace_char_sequence_swaps_literal_substring() {
    let out = run_main(r#"String s = "foo-bar"; System.out.println(s.replace("bar", "baz"));"#);
    assert_eq!(out, vec!["foo-baz"]);
}

#[test]
fn replace_char_unit_swaps_every_matching_code_unit() {
    let out = run_main(r#"String s = "banana"; System.out.println(s.replace('a', 'o'));"#);
    assert_eq!(out, vec!["bonono"]);
}

#[test]
fn replace_char_unit_leaves_string_when_char_absent() {
    let out = run_main(r#"String s = "test"; System.out.println(s.replace('z', 'Z'));"#);
    assert_eq!(out, vec!["test"]);
}

#[test]
fn replace_all_regex_replaces_every_digit_run() {
    let out = run_main(r##"String s = "a12b34"; System.out.println(s.replaceAll("\\d+", "#"));"##);
    assert_eq!(out, vec!["a#b#"]);
}

#[test]
fn replace_all_regex_strips_leading_zeros_from_numbers() {
    let out = run_main(r#"String s = "007"; System.out.println(s.replaceAll("^0+", ""));"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn replace_first_regex_changes_only_initial_digit() {
    let out = run_main(r##"String s = "x1y2z"; System.out.println(s.replaceFirst("\\d", "*"));"##);
    assert_eq!(out, vec!["x*y2z"]);
}

#[test]
fn split_with_limit_zero_drops_trailing_empty_tokens() {
    let out =
        run_main(r#"String[] parts = "a,b,".split(",", 0); System.out.println(parts.length);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn split_with_negative_limit_keeps_all_parts_including_tail() {
    let out = run_main(
        r#"String[] parts = "a,b,c".split(",", -1); System.out.println(parts.length); System.out.println(parts[2]);"#,
    );
    assert_eq!(out, vec!["3", "c"]);
}

#[test]
fn split_with_limit_one_returns_whole_input() {
    let out = run_main(
        r#"String[] parts = "a,b,c".split(",", 1); System.out.println(parts.length); System.out.println(parts[0]);"#,
    );
    assert_eq!(out, vec!["1", "a,b,c"]);
}

#[test]
fn substring_from_begin_index_to_end() {
    let out = run_main(r#"String s = "abcdef"; System.out.println(s.substring(2));"#);
    assert_eq!(out, vec!["cdef"]);
}

#[test]
fn substring_full_range_equals_original_text() {
    let out = run_main(r#"String s = "same"; System.out.println(s.substring(0, 4));"#);
    assert_eq!(out, vec!["same"]);
}

#[test]
fn compare_to_empty_string_is_greater_than_empty() {
    let out = run_main(r#"System.out.println("a".compareTo(""));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn compare_to_shorter_prefix_is_less_than_extension() {
    let out = run_main(r#"System.out.println("app".compareTo("apple"));"#);
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn is_blank_true_for_empty_string() {
    let out = run_main(r#"System.out.println("".isBlank());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strip_indent_removes_shared_leading_spaces() {
    let out = run_main(
        r#"String block = "    line1\n    line2"; System.out.println(block.stripIndent());"#,
    );
    assert_eq!(out, vec!["line1\nline2"]);
}

#[test]
fn strip_indent_preserves_relative_inner_indentation() {
    let out = run_main(
        r#"String block = "  outer\n    inner"; System.out.println(block.stripIndent());"#,
    );
    assert_eq!(out, vec!["outer\n  inner"]);
}

#[test]
fn formatted_fills_multiple_placeholders_in_order() {
    let out = run_main(r#"System.out.println("%s-%d-%s".formatted("vy", 2, "be"));"#);
    assert_eq!(out, vec!["vy-2-be"]);
}

#[test]
fn formatted_zero_pads_integer_width() {
    let out = run_main(r#"System.out.println("%03d".formatted(7));"#);
    assert_eq!(out, vec!["007"]);
}

#[test]
fn translate_escapes_converts_backslash_n_to_newline() {
    let out =
        run_main(r#"String s = String.translateEscapes("\\n"); System.out.println(s.length());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn translate_escapes_converts_backslash_t_to_tab() {
    let out = run_main(
        r#"String s = String.translateEscapes("\\t"); System.out.println((int) s.charAt(0));"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn translate_escapes_unicode_escape_yields_expected_letter() {
    let out = run_main(r#"String s = String.translateEscapes("\\u0041"); System.out.println(s);"#);
    assert_eq!(out, vec!["A"]);
}

#[test]
fn value_of_long_formats_64_bit_integer() {
    let out = run_main(r#"System.out.println(String.valueOf(9223372036854775807L));"#);
    assert_eq!(out, vec!["9223372036854775807"]);
}

#[test]
fn value_of_float_formats_decimal_number() {
    let out = run_main(r#"System.out.println(String.valueOf(2.5f));"#);
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn value_of_char_array_builds_entire_sequence() {
    let out = run_main(
        r#"char[] data = {'j', 'a', 'v', 'a'}; System.out.println(String.valueOf(data));"#,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn value_of_char_array_slice_honors_offset_and_count() {
    let out = run_main(
        r#"char[] data = {'a', 'b', 'c', 'd'}; System.out.println(String.valueOf(data, 1, 2));"#,
    );
    assert_eq!(out, vec!["bc"]);
}

#[test]
fn value_of_object_delegates_to_to_string() {
    let out = run_in_main(
        r#"System.out.println(String.valueOf(new Box("payload")));"#,
        r#"static class Box { String v; Box(String v) { this.v = v; } public String toString() { return "box:" + v; } }"#,
    );
    assert_eq!(out, vec!["box:payload"]);
}

#[test]
fn copy_value_of_char_array_duplicates_content() {
    let out =
        run_main(r#"char[] src = {'x', 'y', 'z'}; System.out.println(String.copyValueOf(src));"#);
    assert_eq!(out, vec!["xyz"]);
}

#[test]
fn copy_value_of_char_array_slice_copies_range_only() {
    let out = run_main(
        r#"char[] src = {'a', 'b', 'c', 'd'}; System.out.println(String.copyValueOf(src, 2, 1));"#,
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn join_varargs_with_no_elements_yields_empty_string() {
    let out = run_main(r#"System.out.println(String.join(",", new String[] {}));"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn join_single_element_has_no_trailing_delimiter() {
    let out = run_main(r#"System.out.println(String.join("-", "solo"));"#);
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn join_iterable_from_arraylist_preserves_order() {
    let out = run_main(
        "java.util.ArrayList<String> items = new java.util.ArrayList<String>(); items.add(\"a\"); items.add(\"b\"); System.out.println(String.join(\"|\", items));",
    );
    assert_eq!(out, vec!["a|b"]);
}
