//! parseInt, parseFloat, Number — parsing edge cases and radix behavior.

crate::js_cases! {
    parseint_decimal_string => {
        r#"console.log(parseInt("42"));"#,
        ["42"]
    };

    parseint_stops_at_non_digit => {
        r#"console.log(parseInt("42px"));"#,
        ["42"]
    };

    parseint_hex_with_radix_16 => {
        r#"console.log(parseInt("ff",16));"#,
        ["255"]
    };

    parseint_binary_with_radix_2 => {
        r#"console.log(parseInt("1010",2));"#,
        ["10"]
    };

    parseint_octal_with_radix_8 => {
        r#"console.log(parseInt("17",8));"#,
        ["15"]
    };

    parseint_leading_whitespace => {
        r#"console.log(parseInt("   7"));"#,
        ["7"]
    };

    parseint_negative_string => {
        r#"console.log(parseInt("-12"));"#,
        ["-12"]
    };

    parseint_empty_string_nan => {
        r#"console.log(Number.isNaN(parseInt("")));"#,
        ["true"]
    };

    parseint_non_string_coerces => {
        r#"console.log(parseInt(42.9));"#,
        ["42"]
    };

    parsefloat_decimal => {
        r#"console.log(parseFloat("3.14"));"#,
        ["3.14"]
    };

    parsefloat_scientific_notation => {
        r#"console.log(parseFloat("1e2"));"#,
        ["100"]
    };

    parsefloat_stops_at_invalid => {
        r#"console.log(parseFloat("3.14abc"));"#,
        ["3.14"]
    };

    parsefloat_negative => {
        r#"console.log(parseFloat("-0.5"));"#,
        ["-0.5"]
    };

    parsefloat_empty_nan => {
        r#"console.log(Number.isNaN(parseFloat("")));"#,
        ["true"]
    };

    parsefloat_whitespace_trim => {
        r#"console.log(parseFloat("  2.5  "));"#,
        ["2.5"]
    };

    number_from_decimal_string => {
        r#"console.log(Number("12.5"));"#,
        ["12.5"]
    };

    number_from_hex_string_prefix => {
        r#"console.log(Number("0x1F"));"#,
        ["31"]
    };

    // Node-verified: numeric separators are LITERAL-only syntax — a
    // runtime string "1_000" is not a StringNumericLiteral → NaN (§7.1.4.1).
    number_from_string_with_underscores_separators => {
        r#"console.log(Number("1_000"));"#,
        ["NaN"]
    };

    number_isfinite_on_infinity => {
        r#"console.log(Number.isFinite(Infinity));"#,
        ["false"]
    };

    number_isfinite_on_integer => {
        r#"console.log(Number.isFinite(42));"#,
        ["true"]
    };

    number_isinteger_on_float => {
        r#"console.log(Number.isInteger(3.14));"#,
        ["false"]
    };

    number_parsefloat_alias => {
        r#"console.log(Number.parseFloat("2.5"));"#,
        ["2.5"]
    };

    number_parseint_alias => {
        r#"console.log(Number.parseInt("10",10));"#,
        ["10"]
    };

    parseint_with_radix_zero_treats_as_decimal => {
        r#"console.log(parseInt("11",0));"#,
        ["11"]
    };

    parseint_hex_without_prefix_radix_16 => {
        r#"console.log(parseInt("1a",16));"#,
        ["26"]
    };

    parsefloat_plus_sign => {
        r#"console.log(parseFloat("+3"));"#,
        ["3"]
    };

    parseint_plus_sign => {
        r#"console.log(parseInt("+8"));"#,
        ["8"]
    };

    number_negative_zero_preserved => {
        r#"console.log(1/Number("-0")<0);"#,
        ["true"]
    };

    number_max_value_finite => {
        r#"console.log(Number.isFinite(Number.MAX_VALUE));"#,
        ["true"]
    };

    number_min_value_positive => {
        r#"console.log(Number.MIN_VALUE>0);"#,
        ["true"]
    };

    number_epsilon_value => {
        r#"console.log(Number.EPSILON>0);"#,
        ["true"]
    };

    // Node-verified: the value exceeds 2^53, so the f64 result rounds to
    // 9007199254741000 — which IS the precision truncation the test name
    // describes.
    parseint_very_large_string_truncates_precision => {
        r#"console.log(parseInt("9007199254740999"));"#,
        ["9007199254741000"]
    };

    parsefloat_negative_infinity => {
        r#"console.log(parseFloat("-Infinity"));"#,
        ["-Infinity"]
    };

    number_from_object_valueof => {
        r#"console.log(Number({valueOf(){return 9;}}));"#,
        ["9"]
    };

    number_from_object_toprimitive_string => {
        r#"console.log(Number({toString(){return "4.5";}}));"#,
        ["4.5"]
    };

    parseint_unicode_digit_stops => {
        r#"console.log(parseInt("１２３"));"#,
        ["NaN"]
    };

    parsefloat_binary_string_partial => {
        r#"console.log(parseFloat("0b10"));"#,
        ["0"]
    };

    parseint_base36_lowercase => {
        r#"console.log(parseInt("z",36));"#,
        ["35"]
    };

    parseint_invalid_radix_nan => {
        r#"console.log(Number.isNaN(parseInt("10",37)));"#,
        ["true"]
    };

    number_isnan_on_string_nan => {
        r#"console.log(Number.isNaN(Number("NaN")));"#,
        ["true"]
    };

    parsefloat_hex_not_supported => {
        r#"console.log(parseFloat("0x10"));"#,
        ["0"]
    };

    number_safe_integer_check => {
        r#"console.log(Number.isSafeInteger(9007199254740991));"#,
        ["true"]
    };

    number_unsafe_integer_check => {
        r#"console.log(Number.isSafeInteger(9007199254740992));"#,
        ["false"]
    };

    parseint_trailing_dot_stops_parse => {
        r#"console.log(parseInt("10.9"));"#,
        ["10"]
    };

    parsefloat_leading_dot_decimal => {
        r#"console.log(parseFloat(".5"));"#,
        ["0.5"]
    };

    number_from_date_object => {
        r#"console.log(Number(new Date(0)));"#,
        ["0"]
    };

    parseint_auto_hex_legacy_string => {
        r#"console.log(parseInt("0x10"));"#,
        ["16"]
    };
}
