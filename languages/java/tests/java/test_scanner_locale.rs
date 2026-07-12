use crate::helpers::run_main;

#[test]
fn scanner_use_locale_france_parses_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,5"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5"]);
}

#[test]
fn scanner_use_locale_germany_parses_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("2,75"); sc.useLocale(java.util.Locale.GERMANY); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["2.75"]);
}

#[test]
fn scanner_use_locale_us_parses_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("3.14"); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn scanner_use_locale_uk_parses_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("9.9"); sc.useLocale(java.util.Locale.UK); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["9.9"]);
}

#[test]
fn scanner_use_locale_france_parses_negative_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("-4,25"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["-4.25"]);
}

#[test]
fn scanner_use_locale_us_parses_negative_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("-4.25"); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["-4.25"]);
}

#[test]
fn scanner_use_locale_france_next_float_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,25"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextFloat());"#,
    );
    assert_eq!(out, vec!["1.25"]);
}

#[test]
fn scanner_use_locale_us_next_float_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1.25"); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextFloat());"#,
    );
    assert_eq!(out, vec!["1.25"]);
}

#[test]
fn scanner_has_next_double_true_for_france_comma_input() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("7,5 more"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.hasNextDouble());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_double_false_for_non_numeric_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("hello"); sc.useLocale(java.util.Locale.US); System.out.println(sc.hasNextDouble());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_has_next_double_true_for_us_dot_input() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("7.5 tail"); sc.useLocale(java.util.Locale.US); System.out.println(sc.hasNextDouble());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_use_radix_sixteen_parses_hex_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("ff"); sc.useRadix(16); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["255"]);
}

#[test]
fn scanner_use_radix_eight_parses_octal_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("77"); sc.useRadix(8); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["63"]);
}

#[test]
fn scanner_use_radix_two_parses_binary_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1010"); sc.useRadix(2); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn scanner_use_radix_sixteen_parses_uppercase_hex() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("AB"); sc.useRadix(16); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["171"]);
}

#[test]
fn scanner_use_radix_ten_is_default_for_decimal_int() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("42"); sc.useRadix(10); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn scanner_has_next_int_true_for_hex_with_radix_sixteen() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1a rest"); sc.useRadix(16); System.out.println(sc.hasNextInt());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_int_false_for_hex_token_with_decimal_radix() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1a"); sc.useRadix(10); System.out.println(sc.hasNextInt());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_use_radix_sixteen_next_long() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("deadbeef"); sc.useRadix(16); System.out.println(sc.nextLong());"#,
    );
    assert_eq!(out, vec!["3735928559"]);
}

#[test]
fn scanner_use_radix_and_locale_together_keep_independent_behavior() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("ff 1,5"); sc.useRadix(16); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextInt()); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["255", "1.5"]);
}

#[test]
fn scanner_use_delimiter_semicolon_with_france_locale_double() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,5;2,0"); sc.useDelimiter(";"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble()); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5", "2.0"]);
}

#[test]
fn scanner_use_delimiter_comma_with_us_locale_doubles() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1.5,2.0"); sc.useDelimiter(","); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextDouble()); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5", "2.0"]);
}

#[test]
fn scanner_use_locale_after_skipping_prefix() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("value=1,75"); sc.skip("[^0-9,-]+"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.75"]);
}

#[test]
fn scanner_find_in_line_with_us_locale_double() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("rate=3.5 done"); sc.useLocale(java.util.Locale.US); System.out.println(sc.findInLine("[0-9.]+"));"#,
    );
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn scanner_find_in_line_with_france_locale_comma_double() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("rate=3,5 done"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.findInLine("[0-9,]+"));"#,
    );
    assert_eq!(out, vec!["3,5"]);
}

#[test]
fn scanner_next_after_use_locale_switch_from_france_to_us() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,5 2.5"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble()); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5", "2.5"]);
}

#[test]
fn scanner_use_radix_sixteen_then_ten() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("10 10"); sc.useRadix(16); System.out.println(sc.nextInt()); sc.useRadix(10); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["16", "10"]);
}

#[test]
fn scanner_has_next_big_decimal_true_for_us_input() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("12.34 tail"); sc.useLocale(java.util.Locale.US); System.out.println(sc.hasNextBigDecimal());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_next_big_decimal_with_france_locale() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("12,34"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextBigDecimal());"#,
    );
    assert_eq!(out, vec!["12.34"]);
}

#[test]
fn scanner_next_big_decimal_with_us_locale() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("12.34"); sc.useLocale(java.util.Locale.US); System.out.println(sc.nextBigDecimal());"#,
    );
    assert_eq!(out, vec!["12.34"]);
}

#[test]
fn scanner_use_locale_italy_parses_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("6,6"); sc.useLocale(java.util.Locale.ITALY); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["6.6"]);
}

#[test]
fn scanner_use_locale_germany_plain_comma_decimal_without_grouping() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1234,5"); sc.useLocale(java.util.Locale.GERMANY); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1234.5"]);
}

#[test]
fn scanner_use_delimiter_pipe_with_radix_sixteen() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("a|b|c"); sc.useDelimiter("\\|"); sc.useRadix(16); System.out.println(sc.next()); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["a", "11"]);
}

#[test]
fn scanner_has_next_with_us_locale_after_double() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1.0 done"); sc.useLocale(java.util.Locale.US); sc.nextDouble(); System.out.println(sc.hasNext());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_double_false_after_consuming_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1.0"); sc.useLocale(java.util.Locale.US); sc.nextDouble(); System.out.println(sc.hasNextDouble());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_use_locale_japan_parses_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("8.8"); sc.useLocale(java.util.Locale.JAPAN); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["8.8"]);
}

#[test]
fn scanner_use_radix_sixteen_negative_hex() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("-10"); sc.useRadix(16); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["-16"]);
}

#[test]
fn scanner_mixed_int_radix_ten_and_hex_word() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("10 dead"); sc.useRadix(10); System.out.println(sc.nextInt()); sc.useRadix(16); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["10", "dead"]);
}

#[test]
fn scanner_use_delimiter_whitespace_and_locale_france() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("  1,5   2,0  "); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble()); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5", "2.0"]);
}

#[test]
fn scanner_next_line_then_double_with_us_locale() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("header\n1.5"); sc.useLocale(java.util.Locale.US); sc.nextLine(); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.5"]);
}

#[test]
fn scanner_has_next_long_true_for_decimal_radix() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("9223372036854775807 rest"); sc.useRadix(10); System.out.println(sc.hasNextLong());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_use_radix_sixteen_has_next_long() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("7fffffffffffffff"); sc.useRadix(16); System.out.println(sc.hasNextLong());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_use_locale_canada_french_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("5,5"); sc.useLocale(java.util.Locale.CANADA_FRENCH); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["5.5"]);
}

#[test]
fn scanner_use_locale_canada_english_dot_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("5.5"); sc.useLocale(java.util.Locale.CANADA); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["5.5"]);
}

#[test]
fn scanner_use_delimiter_colon_with_france_locale() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,25:2,50"); sc.useDelimiter(":"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble()); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["1.25", "2.5"]);
}

#[test]
fn scanner_find_within_horizon_us_double() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("prefix 4.5 suffix"); sc.useLocale(java.util.Locale.US); System.out.println(sc.findWithinHorizon("[0-9.]+", 0));"#,
    );
    assert_eq!(out, vec!["4.5"]);
}

#[test]
fn scanner_skip_us_grouping_not_consumed_as_number() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1,234.5"); sc.useLocale(java.util.Locale.US); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["1,234.5"]);
}

#[test]
fn scanner_use_radix_sixteen_zero() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("0"); sc.useRadix(16); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn scanner_use_locale_france_zero_comma_decimal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("0,0"); sc.useLocale(java.util.Locale.FRANCE); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["0.0"]);
}

#[test]
fn scanner_use_delimiter_custom_with_radix_and_locale() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1.5|ff"); sc.useDelimiter("\\|"); sc.useLocale(java.util.Locale.US); sc.useRadix(16); System.out.println(sc.nextDouble()); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["1.5", "255"]);
}
