crate::js_cases! {
    string_substr_reads_length_from_start_index => {
        r#"
console.log("abcdef".substr(1, 3));
"#,
        ["bcd"]
    };

    string_substr_without_length_runs_to_end => {
        r#"
console.log("abcdef".substr(2));
"#,
        ["cdef"]
    };

    string_substr_negative_start_counts_from_end => {
        r#"
console.log("abcdef".substr(-2));
"#,
        ["ef"]
    };

    string_substr_zero_length_returns_empty_string => {
        r#"
console.log("abcdef".substr(1, 0) === "");
"#,
        ["true"]
    };

    string_startswith_accepts_position_argument => {
        r#"
const s = "banana";
console.log(s.startsWith("na", 2));
console.log(s.startsWith("ba", 2));
"#,
        ["true", "false"]
    };

    string_startswith_position_past_length_is_false => {
        r#"
console.log("abc".startsWith("a", 10));
"#,
        ["false"]
    };

    string_endswith_accepts_length_argument => {
        r#"
const s = "banana";
console.log(s.endsWith("ban", 3));
console.log(s.endsWith("nana", 6));
"#,
        ["true", "true"]
    };

    string_endswith_shorter_length_can_make_match_fail => {
        r#"
console.log("banana".endsWith("nana", 5));
"#,
        ["false"]
    };

    string_includes_empty_string_is_true_even_at_clamped_end => {
        r#"
console.log("abc".includes("", 3));
console.log("abc".includes("", 99));
"#,
        ["true", "true"]
    };

    string_includes_with_regexp_throws_type_error => {
        r#"
try {
  console.log("abc".includes(/a/));
} catch (error) {
  console.log(error instanceof TypeError);
}
"#,
        ["true"]
    };

    string_charcodeat_out_of_range_returns_nan => {
        r#"
console.log(Number.isNaN("abc".charCodeAt(99)));
"#,
        ["true"]
    };

    string_charcodeat_reads_utf16_surrogate_units => {
        r#"
const s = "😀";
console.log(s.charCodeAt(0));
console.log(s.charCodeAt(1));
"#,
        ["55357", "56832"]
    };

    string_codepointat_reads_full_astral_code_point => {
        r#"
console.log("😀".codePointAt(0));
"#,
        ["128512"]
    };

    string_from_code_point_emits_astral_symbol => {
        r#"
const s = String.fromCodePoint(0x1F600);
console.log(s);
console.log(s.length);
"#,
        ["😀", "2"]
    };

    string_normalize_nfd_expands_composed_character => {
        r#"
const s = "é".normalize("NFD");
console.log(s.length);
console.log(s === "e\u0301");
"#,
        ["2", "true"]
    };

    string_normalize_nfkc_folds_compatibility_character => {
        r#"
console.log("Ａ".normalize("NFKC"));
"#,
        ["A"]
    };

    string_substring_negative_arguments_are_clamped_to_zero => {
        r#"
console.log("abcdef".substring(-2, 2));
"#,
        ["ab"]
    };

    string_substring_without_end_uses_rest_of_string => {
        r#"
console.log("abcdef".substring(2));
"#,
        ["cdef"]
    };

    string_repeat_with_fractional_count_uses_integer_part => {
        r#"
console.log("ab".repeat(2.7));
"#,
        ["abab"]
    };

    string_pad_start_truncates_long_filler => {
        r#"
console.log("5".padStart(4, "xyz"));
"#,
        ["xyz5"]
    };

    string_pad_end_truncates_long_filler => {
        r#"
console.log("5".padEnd(4, "xyz"));
"#,
        ["5xyz"]
    };

    string_split_with_empty_separator_returns_code_units => {
        r#"
const parts = "abc".split("");
console.log(parts.length);
console.log(parts.join(","));
"#,
        ["3", "a,b,c"]
    };

    string_split_with_undefined_separator_returns_whole_string => {
        r#"
const parts = "a,b,c".split(undefined);
console.log(parts.length);
console.log(parts[0]);
"#,
        ["1", "a,b,c"]
    };

    string_from_char_code_can_build_multiple_characters => {
        r#"
console.log(String.fromCharCode(72, 105, 33));
"#,
        ["Hi!"]
    };
}