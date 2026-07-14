use crate::helpers::run_main;

#[test]
fn formatter_france_locale_comma_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%.2f", 1.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,50"]);
}

#[test]
fn formatter_germany_locale_comma_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%.2f", 2.25); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["2,25"]);
}

#[test]
fn formatter_italy_locale_comma_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.ITALY); fmt.format("%.1f", 3.14); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["3,1"]);
}

#[test]
fn formatter_us_locale_dot_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%.2f", 9.9); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["9.90"]);
}

#[test]
fn formatter_uk_locale_dot_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.UK); fmt.format("%.2f", 4.4); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["4.40"]);
}

#[test]
fn formatter_japan_locale_dot_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.JAPAN); fmt.format("%.2f", 7.7); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["7.70"]);
}

#[test]
fn formatter_us_locale_comma_groups_large_integer() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,d", 1234567); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,234,567"]);
}

#[test]
fn formatter_germany_locale_dot_groups_large_number() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%,.2f", 1234.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1.234,50"]);
}

#[test]
fn formatter_france_locale_space_groups_large_number() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%,.2f", 1234.56); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1\u{202f}234,56"]);
}

#[test]
fn formatter_us_locale_negative_grouped_integer() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,d", -98765); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["-98,765"]);
}

#[test]
fn formatter_france_locale_negative_decimal() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%.2f", -12.3); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["-12,30"]);
}

#[test]
fn formatter_germany_locale_negative_grouped_decimal() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%,.2f", -2500.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["-2.500,50"]);
}

#[test]
fn formatter_locale_method_returns_constructor_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); System.out.println(fmt.locale().equals(java.util.Locale.FRANCE));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_us_locale_method_matches_us_constant() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); System.out.println(fmt.locale().equals(java.util.Locale.US));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_default_locale_is_not_null() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(); System.out.println(fmt.locale() != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_us_currency_argument_via_number_format() {
    let out = run_main(
        r#"java.text.NumberFormat nf = java.text.NumberFormat.getCurrencyInstance(java.util.Locale.US); java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%s", nf.format(19.99)); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["$19.99"]);
}

#[test]
fn formatter_france_currency_argument_via_number_format() {
    let out = run_main(
        r#"java.text.NumberFormat nf = java.text.NumberFormat.getCurrencyInstance(java.util.Locale.FRANCE); java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%s", nf.format(10.0)); System.out.println(fmt.toString().contains("10"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_germany_currency_argument_via_number_format() {
    let out = run_main(
        r#"java.text.NumberFormat nf = java.text.NumberFormat.getCurrencyInstance(java.util.Locale.GERMANY); java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%s", nf.format(5.5)); System.out.println(fmt.toString().contains("5"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_us_percent_via_number_format() {
    let out = run_main(
        r#"java.text.NumberFormat nf = java.text.NumberFormat.getPercentInstance(java.util.Locale.US); java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%s", nf.format(0.25)); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["25%"]);
}

#[test]
fn formatter_us_chained_format_with_locale_grouping() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,d", 1000).format(" + %.2f", 1.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,000 + 1.50"]);
}

#[test]
fn formatter_france_chained_format_with_locale_decimals() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%.1f", 1.2).format(" / %.1f", 3.4); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,2 / 3,4"]);
}

#[test]
fn formatter_us_plus_flag_on_positive_integer() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%+d", 42); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["+42"]);
}

#[test]
fn formatter_us_zero_pad_flag_unaffected_by_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%05d", 7); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["00007"]);
}

#[test]
fn formatter_us_hex_lowercase_unaffected_by_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%x", 255); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn formatter_france_hex_uppercase_unaffected_by_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%X", 15); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["F"]);
}

#[test]
fn formatter_us_scientific_notation_uses_dot_exponent() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%e", 1000.0); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1.000000e+03"]);
}

#[test]
fn formatter_france_general_format_small_decimal() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%g", 1.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,50000"]);
}

#[test]
fn formatter_us_width_and_precision_with_grouping() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,12.2f", 1234.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["    1,234.50"]);
}

#[test]
fn formatter_germany_width_and_precision_with_grouping() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%,12.2f", 1234.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["    1.234,50"]);
}

#[test]
fn formatter_us_left_align_string_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%-8s|", "vybe"); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["vybe    |"]);
}

#[test]
fn formatter_france_mixed_integer_and_decimal_in_one_format() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%d = %.2f", 2, 2.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["2 = 2,50"]);
}

#[test]
fn formatter_us_mixed_grouped_integer_and_decimal() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,d / %.2f", 1000000, 3.5); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,000,000 / 3.50"]);
}

#[test]
fn formatter_string_builder_appendable_with_us_locale() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); java.util.Formatter fmt = new java.util.Formatter(sb, java.util.Locale.US); fmt.format("%,d", 5000); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["5,000"]);
}

#[test]
fn formatter_string_builder_appendable_with_germany_locale() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); java.util.Formatter fmt = new java.util.Formatter(sb, java.util.Locale.GERMANY); fmt.format("%.2f", 8.8); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["8,80"]);
}

#[test]
fn formatter_italy_grouped_integer() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.ITALY); fmt.format("%,d", 987654); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["987.654"]);
}

#[test]
fn formatter_canada_french_locale_decimal_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.CANADA_FRENCH); fmt.format("%.2f", 6.6); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["6,60"]);
}

#[test]
fn formatter_canada_english_locale_dot_decimal() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.CANADA); fmt.format("%.2f", 6.6); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["6.60"]);
}

#[test]
fn formatter_us_locale_zero_formats_with_fraction() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%.2f", 0.0); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["0.00"]);
}

#[test]
fn formatter_france_locale_zero_formats_with_fraction() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%.2f", 0.0); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["0,00"]);
}

#[test]
fn formatter_us_comma_flag_with_grouping() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%,10d", 42); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["        42"]);
}

#[test]
fn formatter_germany_octal_unaffected_by_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%o", 8); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn formatter_us_boolean_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%b", true); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_france_character_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("%c", 65); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn formatter_us_hash_flag_hex_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%#x", 10); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["0xa"]);
}

#[test]
fn formatter_us_literal_percent_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("rate=%d%%", 50); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["rate=50%"]);
}

#[test]
fn formatter_france_newline_flag_with_locale() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.FRANCE); fmt.format("a%nb"); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn formatter_us_integer_parse_style_grouping_disabled() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); fmt.format("%d", 12345); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn formatter_germany_small_grouped_thousands() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.GERMANY); fmt.format("%,d", 1000); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1.000"]);
}

#[test]
fn formatter_us_out_method_returns_same_formatter_instance() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.US); java.util.Formatter same = fmt.out(); System.out.println(fmt == same);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn formatter_japan_grouped_integer_uses_comma_separator() {
    let out = run_main(
        r#"java.util.Formatter fmt = new java.util.Formatter(java.util.Locale.JAPAN); fmt.format("%,d", 1234567); System.out.println(fmt.toString());"#,
    );
    assert_eq!(out, vec!["1,234,567"]);
}
