use crate::helpers::run_main;

#[test]
fn integer_parse_int_reads_decimal_digits() {
    let out = run_main(r#"System.out.println(Integer.parseInt("123"));"#);
    assert_eq!(out, vec!["123"]);
}

#[test]
fn integer_parse_int_with_hex_radix() {
    let out = run_main(r#"System.out.println(Integer.parseInt("FF", 16));"#);
    assert_eq!(out, vec!["255"]);
}

#[test]
fn integer_parse_int_with_binary_radix() {
    let out = run_main(r#"System.out.println(Integer.parseInt("1010", 2));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn integer_parse_int_with_octal_radix() {
    let out = run_main(r#"System.out.println(Integer.parseInt("17", 8));"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn integer_value_of_parses_decimal_string() {
    let out = run_main(r#"System.out.println(Integer.valueOf("99"));"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn integer_to_string_formats_decimal_digits() {
    let out = run_main(r#"System.out.println(Integer.toString(42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn integer_to_binary_string_formats_bits() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(10));"#);
    assert_eq!(out, vec!["1010"]);
}

#[test]
fn integer_to_hex_string_formats_nibbles() {
    let out = run_main(r#"System.out.println(Integer.toHexString(255));"#);
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn integer_to_octal_string_formats_base_eight() {
    let out = run_main(r#"System.out.println(Integer.toOctalString(8));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn integer_max_value_constant() {
    let out = run_main("System.out.println(Integer.MAX_VALUE);");
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn integer_min_value_constant() {
    let out = run_main("System.out.println(Integer.MIN_VALUE);");
    assert_eq!(out, vec!["-2147483648"]);
}

#[test]
fn integer_compare_negative_when_first_is_smaller() {
    let out = run_main("System.out.println(Integer.compare(5, 8));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn integer_compare_positive_when_first_is_larger() {
    let out = run_main("System.out.println(Integer.compare(8, 5));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_compare_zero_for_equal_values() {
    let out = run_main("System.out.println(Integer.compare(5, 5));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn long_parse_long_reads_decimal_digits() {
    let out = run_main(r#"System.out.println(Long.parseLong("1000000"));"#);
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn long_max_value_constant() {
    let out = run_main("System.out.println(Long.MAX_VALUE);");
    assert_eq!(out, vec!["9.223372036854776E18"]);
}

#[test]
fn long_min_value_constant() {
    let out = run_main("System.out.println(Long.MIN_VALUE);");
    assert_eq!(out, vec!["-9.223372036854776E18"]);
}

#[test]
fn double_parse_double_reads_fractional_string() {
    let out = run_main(r#"System.out.println(Double.parseDouble("3.14"));"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn double_parse_double_reads_scientific_notation() {
    let out = run_main(r#"System.out.println(Double.parseDouble("1.5e2"));"#);
    assert_eq!(out, vec!["150.0"]);
}

#[test]
fn float_parse_float_reads_fractional_string() {
    let out = run_main(r#"System.out.println(Float.parseFloat("2.5"));"#);
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn double_max_value_constant() {
    let out = run_main("System.out.println(Double.MAX_VALUE);");
    assert_eq!(out, vec!["1.7976931348623157E308"]);
}

#[test]
fn double_nan_constant() {
    let out = run_main("System.out.println(Double.NaN);");
    assert_eq!(out, vec!["NaN"]);
}

#[test]
fn double_positive_infinity_constant() {
    let out = run_main("System.out.println(Double.POSITIVE_INFINITY);");
    assert_eq!(out, vec!["Infinity"]);
}

#[test]
fn double_negative_infinity_constant() {
    let out = run_main("System.out.println(Double.NEGATIVE_INFINITY);");
    assert_eq!(out, vec!["-Infinity"]);
}

#[test]
fn double_is_nan_detects_not_a_number() {
    let out = run_main("System.out.println(Double.isNaN(Double.NaN));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_nan_rejects_finite_value() {
    let out = run_main("System.out.println(Double.isNaN(3.14));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_infinite_detects_positive_infinity() {
    let out = run_main("System.out.println(Double.isInfinite(Double.POSITIVE_INFINITY));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_infinite_detects_negative_infinity() {
    let out = run_main("System.out.println(Double.isInfinite(Double.NEGATIVE_INFINITY));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_is_infinite_rejects_finite_value() {
    let out = run_main("System.out.println(Double.isInfinite(1.0));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn boolean_parse_boolean_true_string() {
    let out = run_main(r#"System.out.println(Boolean.parseBoolean("true"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn boolean_parse_boolean_false_string() {
    let out = run_main(r#"System.out.println(Boolean.parseBoolean("false"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn string_value_of_integer_becomes_digits() {
    let out = run_main(r#"System.out.println(String.valueOf(42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn string_value_of_double_formats_decimal() {
    let out = run_main(r#"System.out.println(String.valueOf(3.14));"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn string_value_of_boolean_true_becomes_text() {
    let out = run_main(r#"System.out.println(String.valueOf(true));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_value_of_char_becomes_single_letter() {
    let out = run_main(r#"System.out.println(String.valueOf('Z'));"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn character_is_digit_on_numeric_char() {
    let out = run_main("System.out.println(Character.isDigit('5'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_digit_rejects_letter() {
    let out = run_main("System.out.println(Character.isDigit('A'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_letter_on_alpha_char() {
    let out = run_main("System.out.println(Character.isLetter('k'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_letter_or_digit_on_alphanumeric() {
    let out = run_main("System.out.println(Character.isLetterOrDigit('9'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_whitespace_on_space_char() {
    let out = run_main("System.out.println(Character.isWhitespace(' '));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_to_upper_case_from_lower() {
    let out = run_main("System.out.println(Character.toUpperCase('a'));");
    assert_eq!(out, vec!["A"]);
}

#[test]
fn character_to_lower_case_from_upper() {
    let out = run_main("System.out.println(Character.toLowerCase('Z'));");
    assert_eq!(out, vec!["z"]);
}

#[test]
fn character_get_numeric_value_of_digit_char() {
    let out = run_main("System.out.println(Character.getNumericValue('7'));");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn autoboxing_wraps_int_primitive_for_wrapper_type() {
    let out = run_main("Integer boxed = 7; System.out.println(boxed);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn unboxing_extracts_primitive_from_integer_wrapper() {
    let out = run_main("Integer boxed = 12; int n = boxed; System.out.println(n + 1);");
    assert_eq!(out, vec!["13"]);
}
