crate::js_cases! {
    boolean_empty_string_is_false => {
        r#"console.log(Boolean(""));"#,
        ["false"]
    };

    boolean_non_empty_string_is_true => {
        r#"console.log(Boolean("0"));"#,
        ["true"]
    };

    boolean_zero_is_false => {
        r#"console.log(Boolean(0));"#,
        ["false"]
    };

    boolean_negative_zero_is_false => {
        r#"console.log(Boolean(-0));"#,
        ["false"]
    };

    boolean_nan_is_false => {
        r#"console.log(Boolean(NaN));"#,
        ["false"]
    };

    boolean_null_is_false => {
        r#"console.log(Boolean(null));"#,
        ["false"]
    };

    boolean_undefined_is_false => {
        r#"console.log(Boolean(undefined));"#,
        ["false"]
    };

    boolean_empty_array_is_true => {
        r#"console.log(Boolean([]));"#,
        ["true"]
    };

    boolean_empty_object_is_true => {
        r#"console.log(Boolean({}));"#,
        ["true"]
    };

    number_true_is_one => {
        r#"console.log(Number(true));"#,
        ["1"]
    };

    number_false_is_zero => {
        r#"console.log(Number(false));"#,
        ["0"]
    };

    number_null_is_zero => {
        r#"console.log(Number(null));"#,
        ["0"]
    };

    number_empty_string_is_zero => {
        r#"console.log(Number(""));"#,
        ["0"]
    };

    number_whitespace_string_is_zero => {
        r#"console.log(Number("   "));"#,
        ["0"]
    };

    number_decimal_string_parses => {
        r#"console.log(Number("42"));"#,
        ["42"]
    };

    number_hex_string_parses => {
        r#"console.log(Number("0x10"));"#,
        ["16"]
    };

    number_binary_string_parses => {
        r#"console.log(Number("0b101"));"#,
        ["5"]
    };

    number_octal_string_parses => {
        r#"console.log(Number("0o10"));"#,
        ["8"]
    };

    number_invalid_string_is_nan => {
        r#"console.log(Number.isNaN(Number("abc")));"#,
        ["true"]
    };

    number_undefined_is_nan => {
        r#"console.log(Number.isNaN(Number(undefined)));"#,
        ["true"]
    };

    string_null_is_literal_null => {
        r#"console.log(String(null));"#,
        ["null"]
    };

    string_undefined_is_literal_undefined => {
        r#"console.log(String(undefined));"#,
        ["undefined"]
    };

    string_true_is_literal_true => {
        r#"console.log(String(true));"#,
        ["true"]
    };

    string_number_preserves_decimal_text => {
        r#"console.log(String(3.14));"#,
        ["3.14"]
    };

    string_nan_is_literal_nan => {
        r#"console.log(String(NaN));"#,
        ["NaN"]
    };

    string_array_joins_with_commas => {
        r#"console.log(String([1, 2, 3]));"#,
        ["1,2,3"]
    };

    string_plain_object_uses_object_tag => {
        r#"console.log(String({}));"#,
        ["[object Object]"]
    };

    object_null_returns_object => {
        r#"console.log(typeof Object(null));"#,
        ["object"]
    };

    object_undefined_returns_object => {
        r#"console.log(typeof Object(undefined));"#,
        ["object"]
    };

    string_symbol_uses_symbol_format => {
        r#"console.log(String(Symbol("x")));"#,
        ["Symbol(x)"]
    };

    number_symbol_throws_typeerror => {
        r#"
try {
  Number(Symbol("x"));
  console.log("no error");
} catch (error) {
  console.log(error instanceof TypeError);
}
"#,
        ["true"]
    };

    bigint_number_ten_conversion => {
        r#"console.log(BigInt(10).toString());"#,
        ["10"]
    };
}
