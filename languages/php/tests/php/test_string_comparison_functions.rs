use super::helpers::run_prints;

// ── strncmp ───────────────────────────────────────────────────

#[test]
fn strncmp_equal_prefix() {
    assert_eq!(
        run_prints(r#"<?php echo strncmp("hello world", "hello there", 5); "#),
        vec!["0"]
    );
}

#[test]
fn strncmp_first_less_than_second() {
    assert_eq!(
        run_prints(r#"<?php echo strncmp("abc", "abd", 3) < 0 ? 'less' : 'not less'; "#),
        vec!["less"]
    );
}

#[test]
fn strncmp_first_greater_than_second() {
    assert_eq!(
        run_prints(r#"<?php echo strncmp("abd", "abc", 3) > 0 ? 'greater' : 'not greater'; "#),
        vec!["greater"]
    );
}

#[test]
fn strncmp_length_zero_is_always_equal() {
    assert_eq!(
        run_prints(r#"<?php echo strncmp("xyz", "abc", 0); "#),
        vec!["0"]
    );
}

#[test]
fn strncmp_longer_than_string_compares_full_string() {
    assert_eq!(
        run_prints(r#"<?php echo strncmp("hi", "hi", 100); "#),
        vec!["0"]
    );
}

// ── strncasecmp ───────────────────────────────────────────────

#[test]
fn strncasecmp_case_insensitive_match() {
    assert_eq!(
        run_prints(r#"<?php echo strncasecmp("HELLO", "hello", 5); "#),
        vec!["0"]
    );
}

#[test]
fn strncasecmp_partial_match_case_insensitive() {
    assert_eq!(
        run_prints(r#"<?php echo strncasecmp("FOO bar", "foo baz", 5) === 0 ? 'equal' : 'diff'; "#),
        vec!["equal"]
    );
}

#[test]
fn strncasecmp_differs_beyond_length() {
    assert_eq!(
        run_prints(r#"<?php echo strncasecmp("HELLO", "hello world", 5); "#),
        vec!["0"]
    );
}

// ── substr_compare ────────────────────────────────────────────

#[test]
fn substr_compare_match_at_offset() {
    assert_eq!(
        run_prints(r#"<?php echo substr_compare("hello world", "world", 6); "#),
        vec!["0"]
    );
}

#[test]
fn substr_compare_mismatch_at_offset() {
    assert_eq!(
        run_prints(
            r#"<?php echo substr_compare("hello world", "earth", 6) !== 0 ? 'different' : 'same'; "#
        ),
        vec!["different"]
    );
}

#[test]
fn substr_compare_with_length_limit() {
    assert_eq!(
        run_prints(r#"<?php echo substr_compare("abcdef", "bcd", 1, 3); "#),
        vec!["0"]
    );
}

#[test]
fn substr_compare_case_insensitive() {
    assert_eq!(
        run_prints(r#"<?php echo substr_compare("Hello World", "WORLD", 6, 5, true); "#),
        vec!["0"]
    );
}

#[test]
fn substr_compare_negative_offset_from_end() {
    assert_eq!(
        run_prints(r#"<?php echo substr_compare("hello", "lo", -2); "#),
        vec!["0"]
    );
}

// ── strpbrk ───────────────────────────────────────────────────

#[test]
fn strpbrk_finds_first_matching_char() {
    assert_eq!(
        run_prints(r#"<?php echo strpbrk("hello world", "aeiou"); "#),
        vec!["ello world"]
    );
}

#[test]
fn strpbrk_returns_from_first_match() {
    assert_eq!(
        run_prints(r#"<?php echo strpbrk("abcdef", "df"); "#),
        vec!["def"]
    );
}

#[test]
fn strpbrk_returns_false_when_no_match() {
    assert_eq!(
        run_prints(r#"<?php echo var_export(strpbrk("hello", "xyz"), true); "#),
        vec!["false"]
    );
}

#[test]
fn strpbrk_first_char_matches() {
    assert_eq!(
        run_prints(r#"<?php echo strpbrk("xyz", "x"); "#),
        vec!["xyz"]
    );
}

// ── str_word_count ────────────────────────────────────────────

#[test]
fn str_word_count_basic_count() {
    assert_eq!(
        run_prints(r#"<?php echo str_word_count("Hello World"); "#),
        vec!["2"]
    );
}

#[test]
fn str_word_count_empty_string_is_zero() {
    assert_eq!(run_prints(r#"<?php echo str_word_count(""); "#), vec!["0"]);
}

#[test]
fn str_word_count_hyphenated_word_counted_as_one() {
    assert_eq!(
        run_prints(r#"<?php echo str_word_count("well-known pattern"); "#),
        vec!["2"]
    );
}

#[test]
fn str_word_count_returns_array_mode_1() {
    assert_eq!(
        run_prints(
            r#"<?php
$words = str_word_count("one two three", 1);
echo implode(',', $words);
echo "\n";
"#
        ),
        vec!["one,two,three"]
    );
}

#[test]
fn str_word_count_returns_array_mode_2_with_positions() {
    assert_eq!(
        run_prints(
            r#"<?php
$words = str_word_count("one two", 2);
echo implode(',', array_keys($words));
echo "\n";
"#
        ),
        vec!["0,4"]
    );
}

// ── similar_text ──────────────────────────────────────────────

#[test]
fn similar_text_returns_common_chars() {
    assert_eq!(
        run_prints(r#"<?php echo similar_text("World", "Word"); "#),
        vec!["4"]
    );
}

#[test]
fn similar_text_with_percent() {
    assert_eq!(
        run_prints(
            r#"<?php
similar_text("abc", "abc", $p);
echo $p;
echo "\n";
"#
        ),
        vec!["100"]
    );
}

#[test]
fn similar_text_empty_strings() {
    assert_eq!(
        run_prints(r#"<?php echo similar_text("", ""); "#),
        vec!["0"]
    );
}

// ── levenshtein ───────────────────────────────────────────────

#[test]
fn levenshtein_same_strings() {
    assert_eq!(
        run_prints(r#"<?php echo levenshtein("hello", "hello"); "#),
        vec!["0"]
    );
}

#[test]
fn levenshtein_one_insertion() {
    assert_eq!(
        run_prints(r#"<?php echo levenshtein("kitten", "sitten"); "#),
        vec!["1"]
    );
}

#[test]
fn strcmp_equal_and_ordering_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcmp("abc", "abc");
echo "|";
echo strcmp("abc", "abd") < 0 ? "lt" : "ge";
echo "|";
echo strcmp("abb", "aba") > 0 ? "gt" : "le";
"#
        ),
        vec!["0|lt|gt"]
    );
}

#[test]
fn strcasecmp_casefolding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcasecmp("Hello", "hello");
echo "|";
echo strcasecmp("apple", "Banana") < 0 ? "lt" : "ge";
echo "|";
echo strcasecmp("abc", "ABC");
"#
        ),
        vec!["0|lt|0"]
    );
}

#[test]
fn strnatcmp_numeric_string_order() {
    assert_eq!(
        run_prints(r#"<?php echo strnatcmp('item2', 'item10') < 0 ? 'lt' : 'ge'; "#),
        vec!["lt"]
    );
}

#[test]
fn strnatcasecmp_casefolded_numeric_order() {
    assert_eq!(
        run_prints(r#"<?php echo strnatcasecmp('File2', 'file10') < 0 ? 'lt' : 'ge'; "#),
        vec!["lt"]
    );
}

#[test]
fn strcasecmp_with_symbols() {
    assert_eq!(
        run_prints(r#"<?php echo strcasecmp("A-B", "a-b") === 0 ? 'equal' : 'diff'; "#),
        vec!["equal"]
    );
}

#[test]
fn strpbrk_early_match() {
    assert_eq!(
        run_prints(r#"<?php echo strpbrk('xyz', 'x'); "#),
        vec!["xyz"]
    );
}

#[test]
fn strpbrk_empty_mask_returns_false() {
    assert_eq!(
        run_prints(r#"<?php echo var_export(strpbrk('abc', ''), true); "#),
        vec!["false"]
    );
}

#[test]
fn strcoll_zero_on_equal_strings() {
    assert_eq!(
        run_prints(r#"<?php echo strcoll('hello', 'hello') === 0 ? 'eq' : 'neq'; "#),
        vec!["eq"]
    );
}

#[test]
fn strpos_with_offset_points_into_tail() {
    assert_eq!(
        run_prints(r#"<?php echo strpos('abcabc', 'c', 3); "#),
        vec!["5"]
    );
}

#[test]
fn strcoll_culture_independent_fallback_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
setlocale(LC_COLLATE, "C");
$cmp = strcoll("A", "a");
if ($cmp < 0) {
    echo -1;
} else {
    echo 1;
}
echo "|";
echo strcoll("a", "a");
"#
        ),
        vec!["-1|0"]
    );
}

#[test]
fn strnatcmp_numeric_text_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strnatcmp("file2", "file10");
echo "|";
echo strnatcmp("img9", "img10");
echo "|";
echo strnatcasecmp("abc9", "ABC10");
"#
        ),
        vec!["-1|1|1"]
    );
}

#[test]
fn str_contains_not_found_returns_false_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains("testing", "x") ? "yes" : "no";
echo "|";
var_export(str_starts_with("", "a"));
echo "|";
echo var_export(str_ends_with("abc", "d"), true);
"#
        ),
        vec!["no|false|false"]
    );
}

#[test]
fn strstr_starts_after_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strstr("one@two@three", "@");
echo "|";
echo strstr("one@two@three", "@", true);
echo "|";
echo stristr("One", "o");
"#
        ),
        vec!["@two@three|one|One"]
    );
}

#[test]
fn strcspn_and_strspn_masks_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn("abc123", "0123456789");
echo "|";
echo strcspn("123abc", "0123456789");
echo "|";
echo strspn("abc123", "abc");
"#
        ),
        vec!["3|0|3"]
    );
}

#[test]
fn strtr_simple_map_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strtr("abcdef", "ab", "12");
echo "|";
echo strtr("hello", ["h"=>"H", "l"=>"L"]);
"#
        ),
        vec!["12cdef|HeLLo"]
    );
}
