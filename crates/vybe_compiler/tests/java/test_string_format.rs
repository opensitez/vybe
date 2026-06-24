use crate::helpers::run_main;

#[test]
fn format_percent_s_inserts_string_argument() {
    let out = run_main(r#"System.out.println(String.format("%s", "java"));"#);
    assert_eq!(out, vec!["java"]);
}

#[test]
fn format_percent_d_inserts_integer_argument() {
    let out = run_main(r#"System.out.println(String.format("%d", 42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn format_percent_f_inserts_double_argument() {
    let out = run_main(r#"System.out.println(String.format("%f", 3.5));"#);
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn format_mixed_specifiers_in_order() {
    let out = run_main(r#"System.out.println(String.format("%s=%d", "count", 7));"#);
    assert_eq!(out, vec!["count=7"]);
}

#[test]
fn format_string_int_and_double_together() {
    let out = run_main(r#"System.out.println(String.format("%s %d %f", "pi", 3, 3.14));"#);
    assert_eq!(out, vec!["pi 3 3.14"]);
}

#[test]
fn format_indexed_argument_second_before_first() {
    let out = run_main(r#"System.out.println(String.format("%2$s %1$s", "world", "hello"));"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn format_indexed_integer_reuses_same_argument() {
    let out = run_main(r#"System.out.println(String.format("%1$d plus %1$d", 5));"#);
    assert_eq!(out, vec!["5 plus 5"]);
}

#[test]
fn format_indexed_string_and_int_out_of_order() {
    let out = run_main(r#"System.out.println(String.format("%2$d %1$s", "items", 9));"#);
    assert_eq!(out, vec!["9 items"]);
}

#[test]
fn format_width_right_aligns_integer_in_five_columns() {
    let out = run_main(r#"System.out.println(String.format("%5d", 7));"#);
    assert_eq!(out, vec!["    7"]);
}

#[test]
fn format_width_left_aligns_string_in_six_columns() {
    let out = run_main(r#"System.out.println(String.format("%-6s", "vy"));"#);
    assert_eq!(out, vec!["vy    "]);
}

#[test]
fn format_zero_pad_width_for_integer() {
    let out = run_main(r#"System.out.println(String.format("%03d", 7));"#);
    assert_eq!(out, vec!["007"]);
}

#[test]
fn format_width_with_precision_on_double() {
    let out = run_main(r#"System.out.println(String.format("%8.2f", 3.14159));"#);
    assert_eq!(out, vec!["    3.14"]);
}

#[test]
fn format_literal_percent_escapes_as_double_percent() {
    let out = run_main(r#"System.out.println(String.format("100%%"));"#);
    assert_eq!(out, vec!["100%"]);
}

#[test]
fn format_hex_lowercase_specifier() {
    let out = run_main(r#"System.out.println(String.format("%x", 255));"#);
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn format_hex_uppercase_specifier() {
    let out = run_main(r#"System.out.println(String.format("%X", 255));"#);
    assert_eq!(out, vec!["FF"]);
}

#[test]
fn format_boolean_specifier_lowercase() {
    let out = run_main(r#"System.out.println(String.format("%b", true));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn format_boolean_specifier_for_false() {
    let out = run_main(r#"System.out.println(String.format("%b", false));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn format_character_specifier_from_char_literal() {
    let out = run_main(r#"System.out.println(String.format("%c", 'Z'));"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn format_newline_specifier_inserts_line_break_length() {
    let out = run_main(r#"String s = String.format("a%nb"); System.out.println(s.length());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn formatted_method_on_literal_string() {
    let out = run_main(r#"System.out.println("Hello %s!".formatted("Java"));"#);
    assert_eq!(out, vec!["Hello Java!"]);
}

#[test]
fn formatted_method_with_integer_placeholder() {
    let out = run_main(r#"System.out.println("value=%d".formatted(12));"#);
    assert_eq!(out, vec!["value=12"]);
}

#[test]
fn formatted_method_with_multiple_placeholders() {
    let out = run_main(r#"System.out.println("%s-%d-%s".formatted("vy", 2, "be"));"#);
    assert_eq!(out, vec!["vy-2-be"]);
}

#[test]
fn format_negative_integer_with_sign() {
    let out = run_main(r#"System.out.println(String.format("%d", -15));"#);
    assert_eq!(out, vec!["-15"]);
}

#[test]
fn format_positive_sign_flag_on_integer() {
    let out = run_main(r#"System.out.println(String.format("%+d", 15));"#);
    assert_eq!(out, vec!["+15"]);
}

#[test]
fn format_space_sign_flag_leaves_positive_unsigned() {
    let out = run_main(r#"System.out.println(String.format("% d", 15));"#);
    assert_eq!(out, vec![" 15"]);
}

#[test]
fn format_octal_specifier_for_integer() {
    let out = run_main(r#"System.out.println(String.format("%o", 8));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn format_string_with_width_and_precision() {
    let out = run_main(r#"System.out.println(String.format("%.3s", "abcdef"));"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn format_indexed_width_on_second_argument() {
    let out = run_main(r#"System.out.println(String.format("%2$5s", "skip", "hi"));"#);
    assert_eq!(out, vec!["   hi"]);
}

#[test]
fn format_indexed_zero_pad_on_first_argument() {
    let out = run_main(r#"System.out.println(String.format("%1$04d", 9));"#);
    assert_eq!(out, vec!["0009"]);
}

#[test]
fn format_multiple_literals_between_specifiers() {
    let out = run_main(r#"System.out.println(String.format("[%s] -> %d", "key", 1));"#);
    assert_eq!(out, vec!["[key] -> 1"]);
}

#[test]
fn format_double_scientific_notation_uppercase() {
    let out = run_main(r#"System.out.println(String.format("%E", 1000.0));"#);
    assert_eq!(out, vec!["1.000000E+03"]);
}

#[test]
fn format_double_general_format() {
    let out = run_main(r#"System.out.println(String.format("%g", 3.0));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn format_long_integer_specifier() {
    let out = run_main(r#"System.out.println(String.format("%d", 9223372036854775807L));"#);
    assert_eq!(out, vec!["9223372036854775807"]);
}

#[test]
fn format_float_narrowing_to_two_decimal_places() {
    let out = run_main(r#"System.out.println(String.format("%.2f", 2.5f));"#);
    assert_eq!(out, vec!["2.50"]);
}

#[test]
fn format_empty_string_argument() {
    let out = run_main(r#"System.out.println(String.format("(%s)", ""));"#);
    assert_eq!(out, vec!["()"]);
}

#[test]
fn format_zero_integer_argument() {
    let out = run_main(r#"System.out.println(String.format("n=%d", 0));"#);
    assert_eq!(out, vec!["n=0"]);
}

#[test]
fn format_width_applies_to_hex_output() {
    let out = run_main(r#"System.out.println(String.format("%4x", 15));"#);
    assert_eq!(out, vec!["   f"]);
}

#[test]
fn format_left_align_width_on_integer() {
    let out = run_main(r#"System.out.println(String.format("%-4d", 8));"#);
    assert_eq!(out, vec!["8   "]);
}

#[test]
fn format_three_indexed_arguments_in_custom_order() {
    let out = run_main(r#"System.out.println(String.format("%3$s/%2$d/%1$s", "z", 2, "a"));"#);
    assert_eq!(out, vec!["a/2/z"]);
}

#[test]
fn format_combines_width_precision_and_string_specifier() {
    let out = run_main(r#"System.out.println(String.format("%10.4s", "longword"));"#);
    assert_eq!(out, vec!["    long"]);
}
