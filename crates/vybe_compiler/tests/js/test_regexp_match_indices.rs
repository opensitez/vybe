//! RegExp exec/test, match indices, sticky and global lastIndex behavior.

crate::js_cases! {
    regexp_exec_returns_array_on_match => {
        r#"const r=/a(b)c/; const m=r.exec("abc"); console.log(m[0]);console.log(m[1]);"#,
        ["abc", "b"]
    };

    regexp_exec_returns_null_on_no_match => {
        r#"console.log(/z/.exec("abc")===null);"#,
        ["true"]
    };

    regexp_test_returns_boolean => {
        r#"console.log(/ab/.test("cab"));console.log(/z/.test("abc"));"#,
        ["true", "false"]
    };

    regexp_global_advances_last_index => {
        r#"const r=/a/g; r.exec("aba"); console.log(r.lastIndex);"#,
        ["1"]
    };

    regexp_global_last_index_resets_after_no_match => {
        r#"const r=/a/g; r.exec("aaa"); r.exec("x"); console.log(r.lastIndex);"#,
        ["0"]
    };

    regexp_sticky_does_not_search_before_last_index => {
        r#"const r=/a/y; r.lastIndex=1; console.log(r.exec("aba"));"#,
        ["null"]
    };

    regexp_ignore_case_flag => {
        r#"console.log(/abc/i.test("AbC"));"#,
        ["true"]
    };

    regexp_dotall_dot_matches_newline => {
        r#"console.log(/a.b/s.test("a\nb"));"#,
        ["true"]
    };

    regexp_named_capture_group => {
        r#"const m=/(?<word>[a-z]+)/.exec("hello"); console.log(m.groups.word);"#,
        ["hello"]
    };

    string_match_with_global_returns_all => {
        r#"console.log("a1b2".match(/\d/g).join(","));"#,
        ["1,2"]
    };

    string_match_without_global_returns_first => {
        r#"console.log("a1b2".match(/\d/)[0]);"#,
        ["1"]
    };

    string_matchall_yields_iterator_of_matches => {
        r#"const m=[..."a1b2".matchAll(/\d/g)][0][0]; console.log(m);"#,
        ["1"]
    };

    regexp_replace_with_string_substitution => {
        r#"console.log("a-b".replace(/-/g, "_"));"#,
        ["a_b"]
    };

    regexp_replace_with_function => {
        r#"console.log("ab".replace(/./g,c=>c.toUpperCase()));"#,
        ["AB"]
    };

    regexp_replace_named_groups => {
        r#"console.log("John Doe".replace(/(?<first>\w+) (?<last>\w+)/,"$<last>, $<first>"));"#,
        ["Doe, John"]
    };

    regexp_split_with_limit => {
        r#"console.log("a,b,c".split(/,/,2).join("|"));"#,
        ["a|b"]
    };

    regexp_source_property_escapes => {
        r#"console.log(/\./.source);"#,
        ["\\."]
    };

    regexp_flags_property_order => {
        r#"console.log(/a/gim.flags.includes("g"));"#,
        ["true"]
    };

    regexp_constructor_from_string => {
        r#"console.log(new RegExp("ab+c").test("abbbc"));"#,
        ["true"]
    };

    regexp_unicode_property_escape_letter => {
        r#"console.log(/\p{L}/u.test("A"));"#,
        ["true"]
    };

    regexp_unicode_escape_in_pattern => {
        r#"console.log(/\u0041/.test("A"));"#,
        ["true"]
    };

    regexp_empty_match_at_start => {
        r#"console.log(/^/.exec("abc")[0]);"#,
        [""]
    };

    regexp_quantifier_exact_count => {
        r#"console.log(/a{3}/.test("aaab"));"#,
        ["true"]
    };

    regexp_non_greedy_quantifier => {
        r#"const m=/<.*?>/ .exec("<a><b>"); console.log(m[0]);"#,
        ["<a>"]
    };

    regexp_alternation_matches_second_branch => {
        r#"console.log(/cat|dog/.exec("hotdog")[0]);"#,
        ["dog"]
    };

    regexp_character_class_range => {
        r#"console.log(/[a-z]/.test("m"));"#,
        ["true"]
    };

    regexp_negated_character_class => {
        r#"console.log(/[^0-9]/.test("a"));"#,
        ["true"]
    };

    regexp_lookahead_positive => {
        r#"console.log(/a(?=b)/.exec("ab")[0]);"#,
        ["a"]
    };

    regexp_lookbehind_positive => {
        r#"console.log(/(?<=a)b/.test("ab"));"#,
        ["true"]
    };

    regexp_indices_property_on_match => {
        r#"const r=/(a)(b)/d; const m=r.exec("ab"); console.log(m.indices[0][0]);console.log(m.indices[1][0]);"#,
        ["0", "0"]
    };

    regexp_last_index_writable => {
        r#"const r=/a/g; r.lastIndex=2; console.log(r.exec("aba"));"#,
        // /a/g from lastIndex 2 on "aba": 'a' at index 2 matches (node-verified).
        ["a"]
    };

    regexp_exec_on_object_coercible_string => {
        r#"console.log(/1/.exec({toString(){return "x1y";}})[0]);"#,
        ["1"]
    };

    regexp_to_string_includes_delimiters => {
        r#"console.log(/abc/gi.toString().startsWith("/"));"#,
        ["true"]
    };

    regexp_compile_not_on_instance => {
        r#"const r=/a/; console.log(typeof r.compile);"#,
        ["undefined"]
    };

    regexp_word_boundary_match => {
        r#"console.log(/\bcat\b/.test("a cat!"));"#,
        ["true"]
    };

    regexp_digit_shorthand => {
        r#"console.log(/\d+/.exec("ab12cd")[0]);"#,
        ["12"]
    };

    regexp_whitespace_shorthand => {
        r#"console.log(/\s+/.test("a b"));"#,
        ["true"]
    };

    regexp_backreference => {
        r#"console.log(/(a)\1/.test("aa"));"#,
        ["true"]
    };

    string_search_returns_index => {
        r#"console.log("abcabc".search(/bc/));"#,
        ["1"]
    };

    string_replace_all_with_global_required => {
        r#"console.log("a.a".replaceAll(".","-"));"#,
        ["a-a"]
    };

    regexp_unicode_code_point_flag => {
        r#"console.log(/👍/u.test("👍"));"#,
        ["true"]
    };

    regexp_empty_pattern_matches_empty_string => {
        r#"console.log(/(?:)/.exec("")[0]);"#,
        [""]
    };

    regexp_multiline_dollar_matches_before_end => {
        // Without the `m` flag, `$` matches only at the true end of input, NOT
        // before an interior line terminator: /a$/ fails on "a\nb" (§22.2.2.6).
        r#"console.log(/a$/.test("a\nb"));"#,
        ["false"]
    };

    regexp_multiline_dollar_with_m_flag => {
        r#"console.log(/b$/m.test("a\nb"));"#,
        ["true"]
    };
}
