crate::js_cases! {
    parse_int_radix_thirty_six_uses_alphanumeric_digits => {
        r#"
console.log(parseInt("z", 36));
"#,
        ["35"]
    };

    parse_int_invalid_input_returns_nan => {
        r#"
console.log(Number.isNaN(parseInt("xyz", 10)));
"#,
        ["true"]
    };

    parse_int_hex_prefix_without_radix_is_hex => {
        r#"
console.log(parseInt("0xff"));
"#,
        ["255"]
    };

    parse_int_leading_zero_string_is_decimal_in_modern_js => {
        r#"
console.log(parseInt("010"));
"#,
        ["10"]
    };

    parse_float_leading_whitespace_is_ignored => {
        r#"
console.log(parseFloat("  3.5"));
"#,
        ["3.5"]
    };

    parse_float_infinity_string_returns_infinity => {
        r#"
console.log(parseFloat("Infinity"));
"#,
        ["Infinity"]
    };

    parse_float_invalid_input_returns_nan => {
        r#"
console.log(Number.isNaN(parseFloat("abc")));
"#,
        ["true"]
    };

    isfinite_coerces_string_number => {
        r#"
console.log(isFinite("12"));
"#,
        ["true"]
    };

    isfinite_coerces_null_to_zero => {
        r#"
console.log(isFinite(null));
"#,
        ["true"]
    };

    isfinite_undefined_is_false => {
        r#"
console.log(isFinite(undefined));
"#,
        ["false"]
    };

    isnan_undefined_is_true => {
        r#"
console.log(isNaN(undefined));
"#,
        ["true"]
    };

    isnan_null_is_false => {
        r#"
console.log(isNaN(null));
"#,
        ["false"]
    };

    isnan_blank_string_is_false => {
        r#"
console.log(isNaN("   "));
"#,
        ["false"]
    };

    isnan_non_numeric_string_is_true => {
        r#"
console.log(isNaN("abc"));
"#,
        ["true"]
    };

    encode_uri_preserves_reserved_url_characters => {
        r#"
console.log(encodeURI("https://example.com/a?b=c&d=e#f"));
"#,
        ["https://example.com/a?b=c&d=e#f"]
    };

    encode_uri_encodes_spaces_but_not_slashes => {
        r#"
console.log(encodeURI("https://example.com/a b/c d"));
"#,
        ["https://example.com/a%20b/c%20d"]
    };

    decode_uri_restores_percent_encoded_spaces => {
        r#"
console.log(decodeURI("https://example.com/a%20b/c%20d"));
"#,
        ["https://example.com/a b/c d"]
    };

    encode_uri_component_encodes_slashes_and_question_mark => {
        r#"
console.log(encodeURIComponent("a/b?c=d"));
"#,
        ["a%2Fb%3Fc%3Dd"]
    };

    decode_uri_component_restores_reserved_characters => {
        r#"
console.log(decodeURIComponent("a%2Fb%3Fc%3Dd"));
"#,
        ["a/b?c=d"]
    };

    decode_uri_malformed_sequence_throws_urierror => {
        r#"
try {
  decodeURI("%E0%A4%A");
  console.log("no error");
} catch (error) {
  console.log(error instanceof URIError);
}
"#,
        ["true"]
    };

    decode_uri_component_malformed_sequence_throws_urierror => {
        r#"
try {
  decodeURIComponent("%E0%A4%A");
  console.log("no error");
} catch (error) {
  console.log(error instanceof URIError);
}
"#,
        ["true"]
    };

    global_nan_is_nan => {
        r#"
console.log(Number.isNaN(NaN));
"#,
        ["true"]
    };

    global_infinity_exceeds_large_finite_number => {
        r#"
console.log(Infinity > 1e308);
"#,
        ["true"]
    };

    negative_infinity_is_less_than_large_negative_number => {
        r#"
console.log(-Infinity < -1e308);
"#,
        ["true"]
    };

    undefined_global_has_undefined_type => {
        r#"
console.log(typeof undefined);
"#,
        ["undefined"]
    };

    global_this_refers_to_global_object => {
        r#"
globalThis.__temp = 42;
console.log(__temp);
delete globalThis.__temp;
"#,
        ["42"]
    };

    parse_int_stops_before_exponent_marker => {
        r#"
console.log(parseInt("15e2", 10));
"#,
        ["15"]
    };

    parse_float_reads_sign_prefix => {
        r#"
console.log(parseFloat("-3.5"));
"#,
        ["-3.5"]
    };

    encode_uri_component_encodes_hash_sign => {
        r#"
console.log(encodeURIComponent("a#b"));
"#,
        ["a%23b"]
    };

    decode_uri_component_decodes_plus_sign_literally => {
        r#"
console.log(decodeURIComponent("a%2Bb"));
"#,
        ["a+b"]
    };

    number_parsefloat_alias_matches_global_parsefloat => {
        r#"
console.log(Number.parseFloat("3.14") === parseFloat("3.14"));
"#,
        ["true"]
    };
}