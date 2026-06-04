crate::js_cases! {
    regexp_constructor_from_string_sets_source => {
        r#"
const re = new RegExp("a+");
console.log(re.source);
"#,
        ["a+"]
    };

    regexp_constructor_from_string_sets_flags => {
        r#"
const re = new RegExp("a+", "gi");
console.log(re.flags);
"#,
        ["gi"]
    };

    regexp_constructor_from_regex_copies_source => {
        r#"
const re = new RegExp(/abc/g);
console.log(re.source);
console.log(re.flags);
"#,
        ["abc", "g"]
    };

    regexp_constructor_with_override_flags_replaces_original => {
        r#"
const re = new RegExp(/abc/g, "i");
console.log(re.flags);
"#,
        ["i"]
    };

    regexp_global_ignorecase_multiline_properties => {
        r#"
const re = /abc/gim;
console.log(re.global);
console.log(re.ignoreCase);
console.log(re.multiline);
"#,
        ["true", "true", "true"]
    };

    regexp_exec_basic_match_value_and_index => {
        r#"
const m = /bc/.exec("abcd");
console.log(m[0]);
console.log(m.index);
"#,
        ["bc", "1"]
    };

    regexp_exec_exposes_input_string => {
        r#"
const m = /bc/.exec("abcd");
console.log(m.input);
"#,
        ["abcd"]
    };

    regexp_exec_with_capture_groups_exposes_groups => {
        r#"
const m = /(ab)(cd)/.exec("xabcdz");
console.log(m[1]);
console.log(m[2]);
"#,
        ["ab", "cd"]
    };

    regexp_exec_no_match_returns_null => {
        r#"
console.log(/z/.exec("abcd") === null);
"#,
        ["true"]
    };

    regexp_global_exec_advances_lastindex => {
        r#"
const re = /a/g;
re.exec("a a");
console.log(re.lastIndex);
"#,
        ["1"]
    };

    regexp_global_exec_finds_second_match => {
        r#"
const re = /a/g;
re.exec("a a");
const m = re.exec("a a");
console.log(m.index);
"#,
        ["2"]
    };

    regexp_global_exec_no_more_matches_resets_lastindex => {
        r#"
const re = /a/g;
re.exec("a");
console.log(re.exec("a") === null);
console.log(re.lastIndex);
"#,
        ["true", "0"]
    };

    regexp_test_basic_true_false => {
        r#"
console.log(/ab/.test("xxabyy"));
console.log(/ab/.test("xxyy"));
"#,
        ["true", "false"]
    };

    regexp_global_test_advances_lastindex => {
        r#"
const re = /a/g;
re.test("a a");
console.log(re.lastIndex);
"#,
        ["1"]
    };

    string_match_without_global_returns_match_object => {
        r#"
const m = "abc123".match(/\d+/);
console.log(m[0]);
console.log(m.index);
"#,
        ["123", "3"]
    };

    string_match_with_global_returns_all_matches => {
        r#"
const m = "a1b22c333".match(/\d+/g);
console.log(m.join(","));
"#,
        ["1,22,333"]
    };

    string_search_returns_first_match_index => {
        r#"
console.log("abc123".search(/\d+/));
"#,
        ["3"]
    };

    string_search_returns_negative_one_on_no_match => {
        r#"
console.log("abc".search(/\d+/));
"#,
        ["-1"]
    };

    string_replace_with_string_pattern_replaces_first_match => {
        r#"
console.log("abcabc".replace(/ab/, "XY"));
"#,
        ["XYcabc"]
    };

    string_replace_with_global_pattern_replaces_all_matches => {
        r#"
console.log("abcabc".replace(/ab/g, "XY"));
"#,
        ["XYcXYc"]
    };

    string_replace_with_capture_group_reference => {
        r#"
console.log("2024-07".replace(/(\d{4})-(\d{2})/, "$2/$1"));
"#,
        ["07/2024"]
    };

    string_replace_with_dollar_dollar_emits_literal_dollar => {
        r#"
console.log("abc".replace(/b/, "$$"));
"#,
        ["a$c"]
    };

    string_replace_with_function_receives_match => {
        r#"
console.log("abc123".replace(/\d+/, m => "[" + m + "]"));
"#,
        ["abc[123]"]
    };

    string_split_with_regex_basic => {
        r#"
console.log("a,b,c".split(/,/).join("|"));
"#,
        ["a|b|c"]
    };

    string_split_with_regex_limit => {
        r#"
console.log("a,b,c".split(/,/, 2).join("|"));
"#,
        ["a|b"]
    };

    string_split_with_capturing_group_includes_separator => {
        r#"
console.log("a,b".split(/(,)/).join("|"));
"#,
        ["a|,|b"]
    };

    regexp_alternation_matches_any_branch => {
        r#"
console.log(/cat|dog/.test("dog"));
console.log(/cat|dog/.test("bird"));
"#,
        ["true", "false"]
    };

    regexp_optional_quantifier_allows_absence => {
        r#"
console.log(/colou?r/.test("color"));
console.log(/colou?r/.test("colour"));
"#,
        ["true", "true"]
    };

    regexp_star_quantifier_allows_zero_or_more => {
        r#"
console.log(/ab*c/.test("ac"));
console.log(/ab*c/.test("abbbc"));
"#,
        ["true", "true"]
    };

    regexp_plus_quantifier_requires_one_or_more => {
        r#"
console.log(/ab+c/.test("ac"));
console.log(/ab+c/.test("abbc"));
"#,
        ["false", "true"]
    };

    regexp_anchor_start_and_end_require_full_match => {
        r#"
console.log(/^abc$/.test("abc"));
console.log(/^abc$/.test("xabc"));
"#,
        ["true", "false"]
    };

    regexp_word_boundary_matches_whole_word => {
        r#"
console.log(/\bcat\b/.test("a cat naps"));
console.log(/\bcat\b/.test("concatenate"));
"#,
        ["true", "false"]
    };

    regexp_noncapturing_group_does_not_create_capture_slot => {
        r#"
const m = /(?:ab)(cd)/.exec("abcd");
console.log(m.length);
console.log(m[1]);
"#,
        ["2", "cd"]
    };

    regexp_lazy_quantifier_takes_shortest_match => {
        r#"
const m = /a.+?c/.exec("abbbcbbbc");
console.log(m[0]);
"#,
        ["abbbc"]
    };

    regexp_dot_does_not_match_newline_without_s_flag => {
        r#"
console.log(/a.c/.test("a\nc"));
"#,
        ["false"]
    };

    regexp_multiline_anchor_matches_line_start => {
        r#"
console.log(/^b/m.test("a\nb"));
"#,
        ["true"]
    };

    regexp_tostring_roundtrips_source_and_flags => {
        r#"
console.log(/abc/gi.toString());
"#,
        ["/abc/gi"]
    };

    regexp_lastindex_is_writable_for_global_regex => {
        r#"
const re = /a/g;
re.lastIndex = 2;
console.log(re.lastIndex);
"#,
        ["2"]
    };

    regexp_character_class_matches_digits => {
        r#"
console.log(/[0-9]+/.test("123"));
console.log(/[0-9]+/.test("abc"));
"#,
        ["true", "false"]
    };

    regexp_negated_character_class_excludes_digits => {
        r#"
console.log(/[^0-9]+/.test("abc"));
console.log(/[^0-9]+/.test("123"));
"#,
        ["true", "false"]
    };

    regexp_whitespace_escape_matches_spaces => {
        r#"
console.log(/a\sb/.test("a b"));
console.log(/a\sb/.test("ab"));
"#,
        ["true", "false"]
    };
}
