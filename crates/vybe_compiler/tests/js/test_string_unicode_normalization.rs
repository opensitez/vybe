//! String Unicode normalization, well-formed checks, and code point iteration.

crate::js_cases! {
    string_normalize_nfc_composes_characters => {
        r#"const s="e\u0301"; console.log(s.normalize("NFC")=== "\u00E9");"#,
        ["true"]
    };

    string_normalize_nfd_decomposes => {
        r#"console.log("\u00E9".normalize("NFD").length);"#,
        ["2"]
    };

    string_iswellformed_true_for_ascii => {
        r#"console.log("abc".isWellFormed());"#,
        ["true"]
    };

    string_iswellformed_false_for_lone_surrogate => {
        r#"console.log("\uD800".isWellFormed());"#,
        ["false"]
    };

    string_towellformed_replaces_lone_surrogate => {
        r#"console.log("\uD800".toWellFormed()=== "\uFFFD");"#,
        ["true"]
    };

    string_fromcodepoint_single_code_point => {
        r#"console.log(String.fromCodePoint(65));"#,
        ["A"]
    };

    string_fromcodepoint_multiple_code_points => {
        r#"console.log(String.fromCodePoint(65,66,67));"#,
        ["ABC"]
    };

    string_fromcodepoint_supplementary_plane => {
        r#"console.log(String.fromCodePoint(0x1F600));"#,
        ["😀"]
    };

    string_codepointat_bmp_character => {
        r#"console.log("A".codePointAt(0));"#,
        ["65"]
    };

    string_codepointat_out_of_range_undefined => {
        r#"console.log("a".codePointAt(5));"#,
        ["undefined"]
    };

    string_iterator_yields_code_points_for_emoji => {
        r#"const cp=[..."😀a"][0].codePointAt(0); console.log(cp>0xFFFF);"#,
        ["true"]
    };

    string_length_counts_code_units_not_code_points => {
        r#"console.log("😀".length);"#,
        ["2"]
    };

    string_slice_on_surrogate_pair_preserves_emoji => {
        r#"console.log("x😀y".slice(1,3).length);"#,
        ["2"]
    };

    string_localecompare_returns_number => {
        r#"console.log(typeof "a".localeCompare("b"));"#,
        ["number"]
    };

    string_localecompare_equal_strings_zero => {
        r#"console.log("x".localeCompare("x"));"#,
        ["0"]
    };

    string_normalize_nfkc_compatibility => {
        r#"console.log("\uFB00".normalize("NFKC").length);"#,
        ["2"]
    };

    string_normalize_nfkd_compatibility => {
        r#"console.log("\uFB00".normalize("NFKD")=== "ff");"#,
        ["true"]
    };

    string_includes_unicode_character => {
        r#"console.log("café".includes("é"));"#,
        ["true"]
    };

    string_starts_with_unicode_prefix => {
        r#"console.log("über".startsWith("ü"));"#,
        ["true"]
    };

    string_ends_with_unicode_suffix => {
        r#"console.log("naïve".endsWith("ïve"));"#,
        ["true"]
    };

    string_indexof_unicode_substring => {
        r#"console.log("façade".indexOf("çade"));"#,
        ["2"]
    };

    string_repeat_unicode_character => {
        r#"console.log("😀".repeat(2).length);"#,
        ["4"]
    };

    string_padstart_with_unicode_fill => {
        r#"console.log("x".padStart(3,"★").length);"#,
        ["3"]
    };

    string_padend_with_unicode_fill => {
        r#"console.log("x".padEnd(4,"☆").length);"#,
        ["4"]
    };

    string_trim_includes_unicode_whitespace_nbsp => {
        r#"console.log("\u00A0a\u00A0".trim());"#,
        ["a"]
    };

    string_split_on_unicode_separator => {
        r#"console.log("a→b→c".split("→").join(","));"#,
        ["a,b,c"]
    };

    string_replace_all_unicode_pattern => {
        r#"console.log("a★b★".replaceAll("★","-"));"#,
        ["a-b-"]
    };

    string_match_unicode_property => {
        r#"console.log("Σ".match(/\p{Lu}/u)[0]);"#,
        ["Σ"]
    };

    string_to_uppercase_unicode => {
        r#"console.log("straße".toUpperCase().includes("SS"));"#,
        ["true"]
    };

    string_to_lowercase_unicode => {
        r#"console.log("HELLO".toLowerCase());"#,
        ["hello"]
    };

    string_charat_returns_utf16_unit => {
        r#"console.log("😀".charAt(0).length);"#,
        ["1"]
    };

    string_at_negative_index_returns_last_unit => {
        r#"console.log("abc".at(-1));"#,
        ["c"]
    };

    string_at_on_emoji_index => {
        r#"console.log("x😀y".at(1).length);"#,
        ["2"]
    };

    string_concat_with_unicode => {
        r#"console.log("a".concat("β"));"#,
        ["aβ"]
    };

    string_valueof_returns_primitive_string => {
        r#"console.log(typeof new String("x").valueOf());"#,
        ["string"]
    };

    string_tostring_on_object_wrapper => {
        r#"console.log(Object.prototype.toString.call(new String("hi")));"#,
        ["[object String]"]
    };

    string_iterator_next_on_empty_string => {
        r#"console.log("".iterator().next().done);"#,
        ["true"]
    };

    string_raw_escapes_backslashes => {
        r#"console.log(String.raw`C:\n`);"#,
        ["C:\\n"]
    };

    string_fromcharcode_multiple_units => {
        r#"console.log(String.fromCharCode(72,105));"#,
        ["Hi"]
    };

    string_search_unicode_char => {
        r#"console.log("façade".search("ç"));"#,
        ["2"]
    };

    string_matchall_unicode_global => {
        r#"const n=[..."aαaα".matchAll(/α/g)].length; console.log(n);"#,
        ["2"]
    };

    string_normalize_same_form_idempotent => {
        r#"const s="é"; console.log(s.normalize("NFC")===s.normalize("NFC").normalize("NFC"));"#,
        ["true"]
    };

    string_iswellformed_after_towellformed => {
        r#"console.log("\uD800".toWellFormed().isWellFormed());"#,
        ["true"]
    };

    string_codepointat_on_mixed_string => {
        r#"console.log("A😀".codePointAt(1));"#,
        ["128512"]
    };

    string_split_limit_zero_returns_empty => {
        r#"console.log("a,b".split(",",0).length);"#,
        ["0"]
    };
}
