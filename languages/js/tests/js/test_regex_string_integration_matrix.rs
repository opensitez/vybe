crate::js_cases! {
    string_matchall_with_global_regex_returns_all_matches => {
        r#"
const values = [..."a1b22c333".matchAll(/\d+/g)].map(m => m[0]);
console.log(values.join(","));
"#,
        ["1,22,333"]
    };

    string_matchall_exposes_match_indexes => {
        r#"
const values = [..."a1b22c333".matchAll(/\d+/g)].map(m => m.index);
console.log(values.join(","));
"#,
        ["1,3,6"]
    };

    string_matchall_exposes_capture_groups => {
        r#"
const values = [..."2024-07 2025-08".matchAll(/(\d{4})-(\d{2})/g)].map(m => m[1] + "/" + m[2]);
console.log(values.join(","));
"#,
        ["2024/07,2025/08"]
    };

    string_matchall_without_global_regex_throws_typeerror => {
        r#"
try {
  [..."abc".matchAll(/a/)];
  console.log("no error");
} catch (error) {
  console.log(error instanceof TypeError);
}
"#,
        ["true"]
    };

    string_replace_with_dollar_ampersand_uses_whole_match => {
        r#"
console.log("abc123def".replace(/\d+/, "[$&]"));
"#,
        ["abc[123]def"]
    };

    string_replace_with_dollar_backtick_uses_prefix => {
        r#"
console.log("abc123def".replace(/\d+/, "<$`>"));
"#,
        ["abc<abc>def"]
    };

    string_replace_with_dollar_quote_uses_suffix => {
        r#"
console.log("abc123def".replace(/\d+/, "<$'>"));
"#,
        ["abc<def>def"]
    };

    string_replace_function_receives_capture_groups => {
        r#"
console.log("2024-07".replace(/(\d{4})-(\d{2})/, (_, y, m) => m + "/" + y));
"#,
        ["07/2024"]
    };

    string_replace_global_function_runs_for_each_match => {
        r##"
let count = 0;
"a1b22c333".replace(/\d+/g, () => { count++; return "#"; });
console.log(count);
    "##,
        ["3"]
    };

    string_split_with_regex_no_match_returns_whole_string => {
        r#"
console.log("abc".split(/,/).join(","));
"#,
        ["abc"]
    };

    string_split_with_trailing_separator_keeps_empty_tail => {
        r#"
console.log("a,b,".split(/,/).length);
"#,
        ["3"]
    };

    regexp_exec_result_array_has_expected_length => {
        r#"
const m = /(a)(b)(c)/.exec("abc");
console.log(m.length);
"#,
        ["4"]
    };

    regexp_test_on_global_regex_twice_advances_then_resets => {
        r#"
const re = /a/g;
console.log(re.test("a"));
console.log(re.test("a"));
console.log(re.lastIndex);
"#,
        ["true", "false", "0"]
    };

    string_match_with_no_match_returns_null => {
        r#"
console.log("abc".match(/\d+/) === null);
"#,
        ["true"]
    };

    string_match_global_with_no_match_returns_null => {
        r#"
console.log("abc".match(/\d+/g) === null);
"#,
        ["true"]
    };

    regexp_exec_after_manual_lastindex_starts_from_offset => {
        r#"
const re = /a/g;
re.lastIndex = 2;
const m = re.exec("bbaab");
console.log(m.index);
"#,
        ["2"]
    };

    regexp_source_escapes_forward_slash => {
        r#"
console.log(new RegExp("a/b").source);
"#,
        ["a\\/b"]
    };

    regexp_empty_pattern_matches_empty_string => {
        r#"
console.log(new RegExp("").test(""));
console.log(new RegExp("").test("abc"));
"#,
        ["true", "true"]
    };

    regexp_exec_empty_pattern_returns_zero_index_match => {
        r#"
const m = new RegExp("").exec("abc");
console.log(m[0] === "");
console.log(m.index);
"#,
        ["true", "0"]
    };

    string_replace_on_empty_pattern_prefixes_replacement => {
        r#"
console.log("abc".replace(/(?:)/, "-"));
"#,
        ["-abc"]
    };

    string_split_on_empty_regex_yields_code_units => {
        r#"
console.log("abc".split(/(?:)/).length);
"#,
        ["3"]
    };

    regexp_ignorecase_property_is_true_when_flag_set => {
        r#"
console.log(/abc/i.ignoreCase);
"#,
        ["true"]
    };
}
