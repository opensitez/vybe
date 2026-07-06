//! Builtin prototype method coverage — distinct behaviors only.
crate::js_cases! {
    slice_negative_start => {
        r#"console.log("abcdef".slice(-2));"#,
        ["ef"]
    };

    slice_negative_end => {
        r#"console.log("abcdef".slice(1,-2));"#,
        ["bcd"]
    };

    slice_undefined_end_to_end => {
        r#"console.log("abc".slice(1));"#,
        ["bc"]
    };

    slice_out_of_range_empty => {
        r#"console.log("hi".slice(5));"#,
        [""]
    };

    slice_swapped_negative_normalized => {
        r#"console.log("abc".slice(2,1));"#,
        [""]
    };

    // Node-verified: swapped (1,3) → chars at indices 1..3 = "bc".
    substring_swaps_arguments => {
        r#"console.log("abcdef".substring(3,1));"#,
        ["bc"]
    };

    substring_negative_treated_as_zero => {
        r#"console.log("abc".substring(-1,2));"#,
        ["ab"]
    };

    substr_positive_start_length => {
        r#"console.log("abcdef".substr(2,3));"#,
        ["cde"]
    };

    substr_negative_start_counts_from_end => {
        r#"console.log("abcdef".substr(-3,2));"#,
        ["de"]
    };

    substr_length_beyond_string => {
        r#"console.log("abc".substr(1,10));"#,
        ["bc"]
    };

    indexof_found_at_zero => {
        r#"console.log("abc".indexOf("a"));"#,
        ["0"]
    };

    indexof_not_found => {
        r#"console.log("abc".indexOf("z"));"#,
        ["-1"]
    };

    indexof_with_from_index => {
        r#"console.log("ababa".indexOf("ba",1));"#,
        ["1"]
    };

    indexof_empty_string_at_zero => {
        r#"console.log("abc".indexOf(""));"#,
        ["0"]
    };

    lastindexof_from_end => {
        r#"console.log("ababa".lastIndexOf("ba"));"#,
        ["3"]
    };

    lastindexof_with_from_index => {
        r#"console.log("ababa".lastIndexOf("ba",2));"#,
        ["1"]
    };

    lastindexof_not_found => {
        r#"console.log("abc".lastIndexOf("z"));"#,
        ["-1"]
    };

    includes_found => {
        r#"console.log("hello".includes("ell"));"#,
        ["true"]
    };

    includes_not_found => {
        r#"console.log("hello".includes("xyz"));"#,
        ["false"]
    };

    includes_with_position => {
        r#"console.log("hello".includes("lo",3));"#,
        ["true"]
    };

    includes_empty_string_always => {
        r#"console.log("abc".includes(""));"#,
        ["true"]
    };

    startswith_at_position => {
        r#"console.log("hello world".startsWith("world",6));"#,
        ["true"]
    };

    startswith_false => {
        r#"console.log("hello".startsWith("hi"));"#,
        ["false"]
    };

    endswith_at_position => {
        r#"console.log("hello world".endsWith("hello",5));"#,
        ["true"]
    };

    endswith_false => {
        r#"console.log("hello".endsWith("lo",3));"#,
        ["false"]
    };

    padstart_min_length => {
        r#"console.log("5".padStart(3,"0"));"#,
        ["005"]
    };

    padstart_already_long_enough => {
        r#"console.log("hello".padStart(3,"x"));"#,
        ["hello"]
    };

    padend_min_length => {
        r#"console.log("5".padEnd(4,"0"));"#,
        ["5000"]
    };

    padend_repeat_pad_string => {
        r#"console.log("1".padEnd(5,"23"));"#,
        ["12323"]
    };

    padstart_default_space => {
        r#"console.log("x".padStart(3));"#,
        ["  x"]
    };

    repeat_zero_returns_empty => {
        r#"console.log("abc".repeat(0));"#,
        [""]
    };

    repeat_count_two => {
        r#"console.log("ab".repeat(2));"#,
        ["abab"]
    };

    repeat_throws_on_negative => {
        r#"try{"a".repeat(-1); console.log("ok");}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    trim_removes_both_ends => {
        r#"console.log("  hi  ".trim());"#,
        ["hi"]
    };

    trimstart_only_leading => {
        r#"console.log("  hi  ".trimStart());"#,
        ["hi  "]
    };

    trimend_only_trailing => {
        r#"console.log("  hi  ".trimEnd());"#,
        ["  hi"]
    };

    trim_includes_nbsp => {
        r#"console.log("\u00A0x\u00A0".trim());"#,
        ["x"]
    };

    trimleft_alias_trimstart => {
        r#"console.log("  a".trimLeft());"#,
        ["a"]
    };

    trimright_alias_trimend => {
        r#"console.log("a  ".trimRight());"#,
        ["a"]
    };

    concat_multiple_args => {
        r#"console.log("a".concat("b","c"));"#,
        ["abc"]
    };

    concat_with_numbers => {
        r#"console.log("n=".concat(42));"#,
        ["n=42"]
    };

    charat_in_range => {
        r#"console.log("abc".charAt(1));"#,
        ["b"]
    };

    charat_out_of_range_empty => {
        r#"console.log("a".charAt(5));"#,
        [""]
    };

    charcodeat_in_range => {
        r#"console.log("A".charCodeAt(0));"#,
        ["65"]
    };

    charcodeat_out_of_range_nan => {
        r#"console.log(Number.isNaN("a".charCodeAt(3)));"#,
        ["true"]
    };

    codepointat_bmp => {
        r#"console.log("€".codePointAt(0));"#,
        ["8364"]
    };

    slice_with_explicit_undefined => {
        r#"console.log("abc".slice(1,undefined));"#,
        ["bc"]
    };

    substring_both_undefined => {
        r#"console.log("abc".substring());"#,
        ["abc"]
    };

    substr_zero_length => {
        r#"console.log("abc".substr(1,0));"#,
        [""]
    };

    // Node-verified: "Ab" DOES contain lowercase "b" at index 1; the
    // case-sensitivity concept needs "B" as the needle.
    indexof_case_sensitive => {
        r#"console.log("Ab".indexOf("B"));"#,
        ["-1"]
    };

    // Node-verified: "Ab".includes("b") is true (b at index 1) — the
    // case-sensitive miss needs "B".
    includes_case_sensitive => {
        r#"console.log("Ab".includes("B"));"#,
        ["false"]
    };

    startswith_empty_string => {
        r#"console.log("abc".startsWith(""));"#,
        ["true"]
    };

    endswith_empty_string => {
        r#"console.log("abc".endsWith(""));"#,
        ["true"]
    };

    padstart_empty_pad_uses_space => {
        r#"console.log("x".padStart(2,""));"#,
        [" x"]
    };

    padend_empty_pad_uses_space => {
        r#"console.log("x".padEnd(2,""));"#,
        ["x "]
    };

    repeat_fractional_truncated => {
        r#"console.log("a".repeat(2.9));"#,
        ["aa"]
    };

    trim_only_whitespace => {
        r#"console.log("   ".trim()==="");"#,
        ["true"]
    };

    trimstart_tab_newline => {
        r#"console.log("\t\nab".trimStart());"#,
        ["ab"]
    };

    trimend_tab_newline => {
        r#"console.log("ab\t\n".trimEnd());"#,
        ["ab"]
    };

    slice_on_empty_string => {
        r#"console.log("".slice(0,1));"#,
        [""]
    };

    indexof_from_index_equals_length => {
        r#"console.log("abc".indexOf("c",3));"#,
        ["-1"]
    };

    // Node-verified: a negative fromIndex clamps to 0 (§22.1.3.11), and
    // "a" IS at index 0 → 0.
    lastindexof_from_negative_index => {
        r#"console.log("abcabc".lastIndexOf("a",-2));"#,
        ["0"]
    };

    includes_position_beyond_length => {
        r#"console.log("abc".includes("a",5));"#,
        ["false"]
    };

    startswith_position_beyond_length => {
        r#"console.log("abc".startsWith("a",5));"#,
        ["false"]
    };

    endswith_zero_position => {
        r#"console.log("abc".endsWith("a",0));"#,
        ["false"]
    };

    padstart_unicode_pad_truncated => {
        r#"console.log("x".padStart(4,"ab"));"#,
        ["abax"]
    };

    concat_empty_string => {
        r#"console.log("".concat("a"));"#,
        ["a"]
    };

    charat_negative_index_empty => {
        r#"console.log("abc".charAt(-1));"#,
        [""]
    };

    substr_negative_length_treated_as_zero => {
        r#"console.log("abc".substr(1,-1));"#,
        [""]
    };

    slice_nan_start_treated_as_zero => {
        r#"console.log("abc".slice(NaN,2));"#,
        ["ab"]
    };

    substring_large_indices => {
        r#"console.log("abc".substring(10,20));"#,
        [""]
    };

    indexof_repeated_search => {
        r#"console.log("aaaa".indexOf("aa"));"#,
        ["0"]
    };

    lastindexof_single_char => {
        r#"console.log("abcba".lastIndexOf("b"));"#,
        ["3"]
    };

    includes_at_exact_position => {
        r#"console.log("testing".includes("ing",4));"#,
        ["true"]
    };

    repeat_on_empty_string => {
        r#"console.log("".repeat(5));"#,
        [""]
    };

    trim_preserves_internal_spaces => {
        r#"console.log("  a  b  ".trim());"#,
        ["a  b"]
    };

    padend_already_sufficient => {
        r#"console.log("long".padEnd(2,"."));"#,
        ["long"]
    };

    concat_creates_new_string => {
        r#"const a="a"; const b=a.concat("b"); console.log(a); console.log(b);"#,
        ["a", "ab"]
    };

    slice_does_not_mutate => {
        r#"const s="abc"; s.slice(1); console.log(s);"#,
        ["abc"]
    };

    charcodeat_surrogate_pair => {
        r#"console.log("𝄞".charCodeAt(0).toString(16));"#,
        ["d834"]
    };

    codepointat_surrogate => {
        r#"console.log("𝄞".codePointAt(0).toString(16));"#,
        ["1d11e"]
    };

    starts_with_position_at_boundary => {
        r#"console.log("abc".startsWith("bc",1));"#,
        ["true"]
    };

    ends_with_position_includes_end => {
        r#"console.log("abc".endsWith("c",3));"#,
        ["true"]
    };

    padstart_zero_target_length => {
        r#"console.log("abc".padStart(0,"x"));"#,
        ["abc"]
    };

    indexof_unicode_codepoint => {
        r#"console.log("café".indexOf("é"));"#,
        ["3"]
    };

    includes_unicode => {
        r#"console.log("café".includes("caf"));"#,
        ["true"]
    };

    slice_with_negative_both => {
        r#"console.log("abcdef".slice(-4,-1));"#,
        ["cde"]
    };

    trimstart_preserves_trailing => {
        r#"console.log("  ab  ".trimStart().length);"#,
        ["4"]
    };

    trimend_preserves_leading => {
        r#"console.log("  ab  ".trimEnd().length);"#,
        ["4"]
    };

    repeat_large_count => {
        r#"console.log("x".repeat(3).length);"#,
        ["3"]
    };

    concat_no_args_returns_copy => {
        r#"const s="hi"; console.log(s.concat());"#,
        ["hi"]
    };

    charat_on_empty_string => {
        r#"console.log("".charAt(0));"#,
        [""]
    };

    lastindexof_empty_string => {
        r#"console.log("abc".lastIndexOf(""));"#,
        ["3"]
    };

    padend_multibyte_pad => {
        r#"console.log("1".padEnd(3,"🌟"));"#,
        ["1🌟"]
    };

    substring_single_arg_to_end => {
        r#"console.log("abcdef".substring(3));"#,
        ["def"]
    };

    slice_empty_range => {
        r#"console.log("abc".slice(2,2));"#,
        [""]
    };

    includes_position_at_length_minus_one => {
        r#"console.log("abc".includes("c",2));"#,
        ["true"]
    };

    // Node-verified: position 2 IS "c", so startsWith("c",2) is true;
    // the past-the-end concept needs position 3.
    startswith_at_end_position_false => {
        r#"console.log("abc".startsWith("c",3));"#,
        ["false"]
    };

    endswith_start_position_zero => {
        r#"console.log("abc".endsWith("a",1));"#,
        ["true"]
    };

    trim_vertical_tab => {
        r#"console.log("\vab\v".trim());"#,
        ["ab"]
    };

    indexof_from_index_negative => {
        r#"console.log("abc".indexOf("a",-5));"#,
        ["0"]
    };

    substr_start_beyond_length => {
        r#"console.log("abc".substr(10));"#,
        [""]
    };

    concat_array_coerced => {
        r#"console.log("a".concat([1,2]));"#,
        ["a1,2"]
    };

    repeat_infinity_throws => {
        r#"try{"a".repeat(Infinity); console.log("ok");}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

}
