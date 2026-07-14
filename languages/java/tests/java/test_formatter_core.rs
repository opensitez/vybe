use crate::helpers::run_main;

#[test]
fn string_format_percent_d_formats_integer() {
    let out = run_main(r#"System.out.println(String.format("%d", 42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn string_format_percent_s_formats_string() {
    let out = run_main(r#"System.out.println(String.format("%s", "vybe"));"#);
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn string_format_percent_f_formats_double() {
    let out = run_main(r#"System.out.println(String.format("%f", 3.14));"#);
    assert_eq!(out, vec!["3.140000"]);
}

#[test]
fn string_format_mixed_percent_s_and_percent_d() {
    let out = run_main(r#"System.out.println(String.format("%s=%d", "count", 7));"#);
    assert_eq!(out, vec!["count=7"]);
}

#[test]
fn string_format_labeled_integer_template() {
    let out = run_main(r#"System.out.println(String.format("n=%d", 8));"#);
    assert_eq!(out, vec!["n=8"]);
}

#[test]
fn string_format_multiple_placeholders_in_order() {
    let out = run_main(r#"System.out.println(String.format("%s:%d:%s", "id", 3, "ok"));"#);
    assert_eq!(out, vec!["id:3:ok"]);
}

#[test]
fn string_format_percent_n_inserts_platform_newline() {
    let out = run_main(r#"System.out.println(String.format("a%nb"));"#);
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn string_format_doubled_percent_is_literal_percent_sign() {
    let out = run_main(r#"System.out.println(String.format("100%%"));"#);
    assert_eq!(out, vec!["100%"]);
}

#[test]
fn string_format_width_five_right_aligns_integer() {
    let out = run_main(r#"System.out.println(String.format("%5d", 7));"#);
    assert_eq!(out, vec!["    7"]);
}

#[test]
fn string_format_width_five_left_aligns_with_minus_flag() {
    let out = run_main(r#"System.out.println(String.format("%-5d", 7));"#);
    assert_eq!(out, vec!["7    "]);
}

#[test]
fn string_format_width_on_string_pads_with_spaces() {
    let out = run_main(r#"System.out.println(String.format("%8s", "hi"));"#);
    assert_eq!(out, vec!["      hi"]);
}

#[test]
fn string_format_precision_two_formats_float() {
    let out = run_main(r#"System.out.println(String.format("%.2f", 3.14159));"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn string_format_precision_zero_truncates_fraction() {
    let out = run_main(r#"System.out.println(String.format("%.0f", 9.9));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn string_format_width_and_precision_combined() {
    let out = run_main(r#"System.out.println(String.format("%8.2f", 1.5));"#);
    assert_eq!(out, vec!["    1.50"]);
}

#[test]
fn string_format_zero_pad_flag_for_integers() {
    let out = run_main(r#"System.out.println(String.format("%05d", 42));"#);
    assert_eq!(out, vec!["00042"]);
}

#[test]
fn string_format_plus_flag_prefixes_positive_sign() {
    let out = run_main(r#"System.out.println(String.format("%+d", 5));"#);
    assert_eq!(out, vec!["+5"]);
}

#[test]
fn string_format_plus_flag_on_negative_keeps_minus() {
    let out = run_main(r#"System.out.println(String.format("%+d", -5));"#);
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn string_format_hex_lowercase_placeholder() {
    let out = run_main(r#"System.out.println(String.format("%x", 255));"#);
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn string_format_hex_uppercase_placeholder() {
    let out = run_main(r#"System.out.println(String.format("%X", 255));"#);
    assert_eq!(out, vec!["FF"]);
}

#[test]
fn string_format_octal_placeholder() {
    let out = run_main(r#"System.out.println(String.format("%o", 8));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn string_format_boolean_placeholder() {
    let out = run_main(r#"System.out.println(String.format("%b", true));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_format_character_placeholder() {
    let out = run_main(r#"System.out.println(String.format("%c", 65));"#);
    assert_eq!(out, vec!["A"]);
}

#[test]
fn string_format_negative_integer() {
    let out = run_main(r#"System.out.println(String.format("%d", -99));"#);
    assert_eq!(out, vec!["-99"]);
}

#[test]
fn string_format_zero_integer() {
    let out = run_main(r#"System.out.println(String.format("%d", 0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn string_format_empty_string_argument() {
    let out = run_main(r#"System.out.println(String.format("[%s]", ""));"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn formatter_format_percent_d_appends_integer() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%d", 15); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn formatter_format_percent_s_appends_string() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%s", "java"); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn formatter_format_percent_f_appends_double() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%.1f", 2.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn formatter_chained_format_calls_concatenate() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%s", "a"); fmt.format("-%d", 1); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["a-1"]);
}

#[test]
fn formatter_format_with_width_padding() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%4d", 9); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["   9"]);
}

#[test]
fn formatter_format_with_precision_on_float() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%.3f", 1.2345); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1.235"]);
}

#[test]
fn formatter_format_percent_n_newline() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("x%ny"); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn formatter_format_literal_percent_escape() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("rate=50%%"); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["rate=50%"]);
}

#[test]
fn formatter_format_multiple_arguments() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%s %d", "item", 4); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["item 4"]);
}

#[test]
fn formatter_locale_us_formats_decimal_point() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%.2f", 1.2); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1.20"]);
}

#[test]
fn formatter_out_returns_same_formatter_for_chaining() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); fmt.format("%d", 1).format("%d", 2); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn string_format_scientific_notation_lowercase() {
    let out = run_main(r#"System.out.println(String.format("%e", 1000.0));"#);
    assert_eq!(out, vec!["1.000000e+03"]);
}

#[test]
fn string_format_scientific_notation_uppercase() {
    let out = run_main(r#"System.out.println(String.format("%E", 1000.0));"#);
    assert_eq!(out, vec!["1.000000E+03"]);
}

#[test]
fn string_format_hash_flag_alternate_hex_form() {
    let out = run_main(r#"System.out.println(String.format("%#x", 10));"#);
    assert_eq!(out, vec!["0xa"]);
}

#[test]
fn string_format_comma_grouping_notation_for_large_int() {
    let out = run_main(r#"System.out.println(String.format("%,d", 1234567));"#);
    assert_eq!(out, vec!["1,234,567"]);
}
