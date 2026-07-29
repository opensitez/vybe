crate::js_cases! {
    parse_int_with_hex_radix_reads_hex_digits => {
        r#"
console.log(parseInt("ff", 16));
"#,
        ["255"]
    };

    parse_int_with_binary_radix_reads_binary_digits => {
        r#"
console.log(parseInt("101", 2));
"#,
        ["5"]
    };

    parse_int_with_octal_radix_reads_octal_digits => {
        r#"
console.log(parseInt("17", 8));
"#,
        ["15"]
    };

    parse_int_with_radix_ignores_trailing_characters => {
        r#"
console.log(parseInt("15px", 10));
"#,
        ["15"]
    };

    parse_int_handles_leading_whitespace_and_signs => {
        r#"
console.log(parseInt("  -12", 10));
console.log(parseInt("+12", 10));
"#,
        ["-12", "12"]
    };

    number_parse_int_accepts_prefixed_hex_with_matching_radix => {
        r#"
console.log(Number.parseInt("0xff", 16));
"#,
        ["255"]
    };

    parse_float_reads_exponent_notation => {
        r#"
console.log(parseFloat("1.25e2"));
"#,
        ["125"]
    };

    parse_float_stops_at_second_decimal_point => {
        r#"
console.log(parseFloat("3.14.15"));
"#,
        ["3.14"]
    };

    number_max_value_is_larger_than_one => {
        r#"
console.log(Number.MAX_VALUE > 1);
"#,
        ["true"]
    };

    number_min_value_is_positive => {
        r#"
console.log(Number.MIN_VALUE > 0);
"#,
        ["true"]
    };

    number_positive_and_negative_infinity_constants_match_globals => {
        r#"
console.log(Number.POSITIVE_INFINITY === Infinity);
console.log(Number.NEGATIVE_INFINITY === -Infinity);
"#,
        ["true", "true"]
    };

    number_nan_constant_is_nan => {
        r#"
console.log(Number.isNaN(Number.NaN));
"#,
        ["true"]
    };

    number_is_finite_does_not_coerce_null => {
        r#"
console.log(Number.isFinite(null));
console.log(isFinite(null));
"#,
        ["false", "true"]
    };

    number_to_string_with_base_thirty_six_uses_alphanumeric_digits => {
        r#"
console.log((255).toString(36));
"#,
        ["73"]
    };

    negative_zero_to_string_normalizes_to_plain_zero => {
        r#"
console.log((-0).toString());
"#,
        ["0"]
    };

    number_value_of_returns_primitive_number => {
        r#"
console.log((12.5).valueOf());
"#,
        ["12.5"]
    };

    number_constructor_without_new_converts_string_to_number => {
        r#"
console.log(Number("42"));
"#,
        ["42"]
    };

    number_object_with_new_has_object_type => {
        r#"
console.log(typeof new Number(42));
"#,
        ["object"]
    };

    number_object_value_of_unboxes_wrapped_number => {
        r#"
console.log(new Number(42).valueOf());
"#,
        ["42"]
    };

    number_to_precision_with_small_precision_uses_exponential_notation => {
        r#"
console.log((12345).toPrecision(2));
"#,
        ["1.2e+4"]
    };

    number_to_exponential_without_digits_uses_default_precision => {
        r#"
console.log((12).toExponential());
"#,
        ["1.2e+1"]
    };

    number_is_integer_and_safe_integer_checks => {
        r#"
console.log(Number.isInteger(10));
console.log(Number.isInteger(10.1));
console.log(Number.isSafeInteger(9007199254740991));
console.log(Number.isSafeInteger(9007199254740992));
"#,
        ["true", "false", "true", "false"]
    };

    number_min_max_safe_integer_are_numbers => {
        r#"
console.log(Number.MIN_SAFE_INTEGER);
console.log(Number.MAX_SAFE_INTEGER);
console.log(Number.MIN_SAFE_INTEGER + 2);
"#,
        ["-9007199254740991", "9007199254740991", "-9007199254740989"]
    };

    number_parse_float_handles_unicode_spaces => {
        r#"
console.log(parseFloat("\t  42.5\n"));
console.log(parseFloat("-10.5foo"));
"#,
        ["42.5", "-10.5"]
    };

    parse_int_and_float_empty_string_returns_nan => {
        r#"
console.log(Number.isNaN(parseInt("")) + "|" + Number.isNaN(parseFloat("")));
"#,
        ["true|true"]
    };
}

