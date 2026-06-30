use crate::helpers::run_main;

#[test]
fn decimal_format_grouped_pattern_two_fraction_digits() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); System.out.println(df.format(1234.5));"##,
    );
    assert_eq!(out, vec!["1,234.50"]);
}

#[test]
fn decimal_format_zero_pads_fraction_to_two_places() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); System.out.println(df.format(42));"##,
    );
    assert_eq!(out, vec!["42.00"]);
}

#[test]
fn decimal_format_negative_value_with_grouping() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); System.out.println(df.format(-9876.1));"##,
    );
    assert_eq!(out, vec!["-9,876.10"]);
}

#[test]
fn decimal_format_million_inserts_two_group_separators() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); System.out.println(df.format(1234567.89));"##,
    );
    assert_eq!(out, vec!["1,234,567.89"]);
}

#[test]
fn decimal_format_optional_fraction_strips_trailing_zeros() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.###", sym); System.out.println(df.format(1.2000));"##,
    );
    assert_eq!(out, vec!["1.2"]);
}

#[test]
fn decimal_format_optional_fraction_keeps_significant_digits() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.###", sym); System.out.println(df.format(3.1415));"##,
    );
    assert_eq!(out, vec!["3.142"]);
}

#[test]
fn decimal_format_minimum_integer_digits_zero_pads() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0000", sym); System.out.println(df.format(7));"##,
    );
    assert_eq!(out, vec!["0007"]);
}

#[test]
fn decimal_format_percent_multiplies_and_appends_sign() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0%", sym); System.out.println(df.format(0.25));"##,
    );
    assert_eq!(out, vec!["25%"]);
}

#[test]
fn decimal_format_percent_whole_number() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0%", sym); System.out.println(df.format(1.0));"##,
    );
    assert_eq!(out, vec!["100%"]);
}

#[test]
fn decimal_format_percent_negative_value() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.0%", sym); System.out.println(df.format(-0.125));"##,
    );
    assert_eq!(out, vec!["-12.5%"]);
}

#[test]
fn decimal_format_per_mille_pattern() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.0\u2030", sym); System.out.println(df.format(0.05));"##,
    );
    assert_eq!(out, vec!["50.0\u{2030}"]);
}

#[test]
fn decimal_format_parse_simple_decimal() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0.00", sym); Number n = df.parse("12.34"); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["12.34"]);
}

#[test]
fn decimal_format_parse_grouped_number() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); Number n = df.parse("1,234.56"); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["1234.56"]);
}

#[test]
fn decimal_format_parse_negative_decimal() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0.00", sym); Number n = df.parse("-45.67"); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["-45.67"]);
}

#[test]
fn decimal_format_parse_percent_string() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0%", sym); Number n = df.parse("75%"); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["0.75"]);
}

#[test]
fn decimal_format_format_parse_roundtrip_grouped() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); String s = df.format(98765.4); Number n = df.parse(s); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["98765.4"]);
}

#[test]
fn decimal_format_format_parse_roundtrip_optional_fraction() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.###", sym); String s = df.format(2.5); Number n = df.parse(s); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn decimal_format_format_parse_roundtrip_percent() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.##%", sym); String s = df.format(0.333); Number n = df.parse(s); System.out.println(Math.abs(n.doubleValue() - 0.333) < 0.001);"##,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn decimal_format_set_minimum_fraction_digits_expands_zeros() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0", sym); df.setMinimumFractionDigits(3); System.out.println(df.format(5));"##,
    );
    assert_eq!(out, vec!["5.000"]);
}

#[test]
fn decimal_format_set_maximum_fraction_digits_truncates() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.0000", sym); df.setMaximumFractionDigits(2); System.out.println(df.format(1.9999));"##,
    );
    assert_eq!(out, vec!["2.00"]);
}

#[test]
fn decimal_format_min_and_max_fraction_digits_together() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0", sym); df.setMinimumFractionDigits(2); df.setMaximumFractionDigits(2); System.out.println(df.format(3.1)); System.out.println(df.format(3.14159));"##,
    );
    assert_eq!(out, vec!["3.10", "3.14"]);
}

#[test]
fn decimal_format_get_minimum_fraction_digits_default() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); System.out.println(df.getMinimumFractionDigits());"##,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn decimal_format_get_maximum_fraction_digits_default() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.###", sym); System.out.println(df.getMaximumFractionDigits());"##,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn decimal_format_set_grouping_used_false_removes_commas() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); df.setGroupingUsed(false); System.out.println(df.format(1234567.0));"##,
    );
    assert_eq!(out, vec!["1234567.00"]);
}

#[test]
fn decimal_format_set_grouping_used_true_restores_commas() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); df.setGroupingUsed(false); df.setGroupingUsed(true); System.out.println(df.format(1234567.0));"##,
    );
    assert_eq!(out, vec!["1,234,567.00"]);
}

#[test]
fn decimal_format_is_grouping_used_reflects_setting() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0", sym); System.out.println(df.isGroupingUsed()); df.setGroupingUsed(false); System.out.println(df.isGroupingUsed());"##,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn decimal_format_negative_subpattern_uses_parentheses() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00;(#,##0.00)", sym); System.out.println(df.format(-2500.5));"##,
    );
    assert_eq!(out, vec!["(2,500.50)"]);
}

#[test]
fn decimal_format_positive_subpattern_formats_normally() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00;(#,##0.00)", sym); System.out.println(df.format(2500.5));"##,
    );
    assert_eq!(out, vec!["2,500.50"]);
}

#[test]
fn decimal_format_apply_pattern_changes_output() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0", sym); df.applyPattern("0.000"); System.out.println(df.format(1.2));"##,
    );
    assert_eq!(out, vec!["1.200"]);
}

#[test]
fn decimal_format_to_pattern_returns_applied_pattern() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.###", sym); System.out.println(df.toPattern());"##,
    );
    assert_eq!(out, vec!["0.###"]);
}

#[test]
fn decimal_format_hash_pattern_formats_integer_without_fraction() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#", sym); System.out.println(df.format(9000));"##,
    );
    assert_eq!(out, vec!["9000"]);
}

#[test]
fn decimal_format_parse_integer_from_hash_pattern() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#", sym); Number n = df.parse("9000"); System.out.println(n.longValue());"##,
    );
    assert_eq!(out, vec!["9000"]);
}

#[test]
fn decimal_format_zero_with_hash_only_fraction() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#.##", sym); System.out.println(df.format(0.0));"##,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn decimal_format_currency_pattern_with_symbol() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("\u00a4#,##0.00", sym); System.out.println(df.format(19.5));"##,
    );
    assert_eq!(out, vec!["$19.50"]);
}

#[test]
fn decimal_format_get_currency_instance_formats_us_dollars() {
    let out = run_main(
        r##"java.text.DecimalFormat df = (java.text.DecimalFormat) java.text.NumberFormat.getCurrencyInstance(java.util.Locale.US); System.out.println(df.format(7.25));"##,
    );
    assert_eq!(out, vec!["$7.25"]);
}

#[test]
fn decimal_format_scientific_pattern_uses_exponent() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.000E0", sym); System.out.println(df.format(12345.0));"##,
    );
    assert_eq!(out, vec!["1.235E4"]);
}

#[test]
fn decimal_format_decimal_separator_always_shown() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0.##", sym); df.setDecimalSeparatorAlwaysShown(true); System.out.println(df.format(42));"##,
    );
    assert_eq!(out, vec!["42."]);
}

#[test]
fn decimal_format_get_multiplier_for_percent_is_hundred() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0%", sym); System.out.println(df.getMultiplier());"##,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn decimal_format_set_multiplier_scales_output() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0", sym); df.setMultiplier(1000); System.out.println(df.format(4));"##,
    );
    assert_eq!(out, vec!["4000"]);
}

#[test]
fn decimal_format_format_long_value() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0", sym); System.out.println(df.format(1000000L));"##,
    );
    assert_eq!(out, vec!["1,000,000"]);
}

#[test]
fn decimal_format_parse_long_from_grouped_string() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0", sym); Number n = df.parse("1,000,000"); System.out.println(n.longValue());"##,
    );
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn decimal_format_rounds_half_up_to_two_places() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.00", sym); System.out.println(df.format(2.345));"##,
    );
    assert_eq!(out, vec!["2.35"]);
}

#[test]
fn decimal_format_truncates_when_max_fraction_exceeded() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.00", sym); df.setMaximumFractionDigits(1); System.out.println(df.format(9.99));"##,
    );
    assert_eq!(out, vec!["10.0"]);
}

#[test]
fn decimal_format_parse_returns_double_for_fractional_input() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0.0", sym); Number n = df.parse("3.5"); System.out.println(n instanceof Double);"##,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn decimal_format_clone_produces_independent_formatter() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.00", sym); java.text.DecimalFormat copy = (java.text.DecimalFormat) df.clone(); copy.setGroupingUsed(false); System.out.println(df.format(1000.0)); System.out.println(copy.format(1000.0));"##,
    );
    assert_eq!(out, vec!["1,000.00", "1000.00"]);
}

#[test]
fn decimal_format_equals_same_pattern_and_symbols() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat a = new java.text.DecimalFormat("0.00", sym); java.text.DecimalFormat b = new java.text.DecimalFormat("0.00", sym); System.out.println(a.equals(b));"##,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn decimal_format_parse_leading_plus_sign() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0.00", sym); Number n = df.parse("+88.00"); System.out.println(n.doubleValue());"##,
    );
    assert_eq!(out, vec!["88.0"]);
}

#[test]
fn decimal_format_small_fraction_near_zero() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("0.0000", sym); System.out.println(df.format(0.0001));"##,
    );
    assert_eq!(out, vec!["0.0001"]);
}

#[test]
fn decimal_format_parse_integer_only_rejects_fraction() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("#0", sym); df.setParseIntegerOnly(true); try { df.parse("12.34"); System.out.println("fail"); } catch (java.text.ParseException e) { System.out.println("ok"); }"##,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn decimal_format_fixed_width_before_decimal_point() {
    let out = run_main(
        r##"java.text.DecimalFormatSymbols sym = new java.text.DecimalFormatSymbols(java.util.Locale.US); java.text.DecimalFormat df = new java.text.DecimalFormat("00.0", sym); System.out.println(df.format(5.5));"##,
    );
    assert_eq!(out, vec!["05.5"]);
}
