use crate::helpers::run_main;

#[test]
fn stringbuilder_set_char_at_replaces_middle_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); sb.setCharAt(1, 'X'); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["aXc"]);
}

#[test]
fn stringbuilder_set_char_at_replaces_first_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("java"); sb.setCharAt(0, 'J'); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["Java"]);
}

#[test]
fn stringbuilder_set_char_at_replaces_last_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("end"); sb.setCharAt(2, '!'); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["en!"]);
}

#[test]
fn stringbuilder_set_length_truncates_buffer() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.setLength(3); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn stringbuilder_set_length_zero_clears_content() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("data"); sb.setLength(0); System.out.println(sb.toString()); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["", "0"]);
}

#[test]
fn stringbuilder_set_length_extends_with_null_chars() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ab"); sb.setLength(4); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringbuilder_substring_returns_middle_slice() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("hello"); System.out.println(sb.substring(1, 4));"#,
    );
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn stringbuilder_substring_from_start() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("vybe"); System.out.println(sb.substring(0, 2));"#,
    );
    assert_eq!(out, vec!["vy"]);
}

#[test]
fn stringbuilder_substring_to_end() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("vybe"); System.out.println(sb.substring(2));"#,
    );
    assert_eq!(out, vec!["be"]);
}

#[test]
fn stringbuilder_replace_range_inserts_new_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.replace(2, 5, "XYZ"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abXYZf"]);
}

#[test]
fn stringbuilder_replace_range_at_start() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("old-new"); sb.replace(0, 3, "new"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["new-new"]);
}

#[test]
fn stringbuilder_replace_range_at_end() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("prefix-old"); sb.replace(7, 10, "new"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["prefix-new"]);
}

#[test]
fn stringbuilder_index_of_finds_first_occurrence() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("banana"); System.out.println(sb.indexOf("na"));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringbuilder_index_of_returns_minus_one_when_missing() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.indexOf("z"));"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn stringbuilder_index_of_with_from_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("banana"); System.out.println(sb.indexOf("na", 3));"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringbuilder_last_index_of_finds_final_occurrence() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("banana"); System.out.println(sb.lastIndexOf("na"));"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringbuilder_last_index_of_with_from_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("banana"); System.out.println(sb.lastIndexOf("na", 3));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringbuilder_char_at_reads_middle_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcde"); System.out.println(sb.charAt(2));"#,
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn stringbuilder_char_at_reads_first_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("zap"); System.out.println(sb.charAt(0));"#,
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn stringbuilder_char_at_reads_last_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("zap"); System.out.println(sb.charAt(2));"#,
    );
    assert_eq!(out, vec!["p"]);
}

#[test]
fn stringbuilder_ensure_capacity_grows_tracked_capacity() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.ensureCapacity(64); System.out.println(sb.capacity() >= 64);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringbuilder_ensure_capacity_no_shrink_when_already_large() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(128); sb.ensureCapacity(32); System.out.println(sb.capacity() >= 32);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringbuilder_capacity_reports_initial_default() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); System.out.println(sb.capacity() >= 16);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringbuilder_capacity_with_seed_string() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("hello"); System.out.println(sb.capacity() >= 5);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringbuilder_compare_to_equal_strings() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.compareTo(new StringBuilder("abc")));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_compare_to_less_than() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.compareTo(new StringBuilder("abd")));"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn stringbuilder_compare_to_greater_than() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.compareTo(new StringBuilder("abb")));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_compare_to_case_sensitive() {
    // JLS String.compareTo: difference of the first differing chars —
    // 'A' (65) - 'a' (97) == -32.
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("A"); System.out.println(sb.compareTo(new StringBuilder("a")));"#,
    );
    assert_eq!(out, vec!["-32"]);
}

#[test]
fn stringbuilder_code_point_at_ascii_position() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ABC"); System.out.println(sb.codePointAt(1));"#,
    );
    assert_eq!(out, vec!["66"]);
}

#[test]
fn stringbuilder_code_point_at_emoji_leading_surrogate() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("a\uD83D\uDE00b"); System.out.println(sb.codePointAt(1));"#,
    );
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn stringbuilder_set_char_at_then_compare_to() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("cat"); sb.setCharAt(0, 'b'); System.out.println(sb.compareTo(new StringBuilder("bat")));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_set_length_then_substring() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.setLength(4); System.out.println(sb.substring(1, 3));"#,
    );
    assert_eq!(out, vec!["bc"]);
}

#[test]
fn stringbuilder_replace_then_index_of() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("foo-bar"); sb.replace(0, 3, "baz"); System.out.println(sb.indexOf("baz"));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_index_of_empty_string_is_zero() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("x"); System.out.println(sb.indexOf(""));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_last_index_of_single_char() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ababa"); System.out.println(sb.lastIndexOf("a"));"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringbuilder_char_at_after_set_char_at() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("pin"); sb.setCharAt(1, 'o'); System.out.println(sb.charAt(1));"#,
    );
    assert_eq!(out, vec!["o"]);
}

#[test]
fn stringbuilder_substring_after_replace() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("123456"); sb.replace(2, 4, "XX"); System.out.println(sb.substring(1, 4));"#,
    );
    assert_eq!(out, vec!["2XX"]);
}

#[test]
fn stringbuilder_compare_to_after_mutations() {
    let out = run_main(
        r#"StringBuilder a = new StringBuilder("aa"); a.append("b"); StringBuilder b = new StringBuilder("aab"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_capacity_after_append_growth() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(8); for (int i = 0; i < 20; i++) { sb.append("x"); } System.out.println(sb.length()); System.out.println(sb.capacity() >= sb.length());"#,
    );
    assert_eq!(out, vec!["20", "true"]);
}

#[test]
fn stringbuilder_code_point_count_via_length_and_surrogates() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("a\uD83D\uDE00"); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_replace_zero_width_at_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); sb.replace(1, 1, "-"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["a-bc"]);
}

#[test]
fn stringbuilder_set_length_shorter_then_char_at() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.setLength(2); System.out.println(sb.charAt(1));"#,
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn stringbuilder_index_of_char_sequence_at_end() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("prefix-end"); System.out.println(sb.indexOf("end"));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn stringbuilder_last_index_of_not_found() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.lastIndexOf("z"));"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn stringbuilder_ensure_capacity_then_set_char_at() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("test"); sb.ensureCapacity(100); sb.setCharAt(0, 'T'); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["Test"]);
}
