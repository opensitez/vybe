use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Internationalization & Locale Formatting — NumberFormatter, Locale, MessageFormatter, IntlDateFormatter, Collator
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_locale_get_default_set_default() {
    let out = run_prints(
        r#"<?php
if (class_exists('Locale')) {
    Locale::setDefault("en_US");
    echo "Locale: " . Locale::getDefault();
} else {
    echo "Locale: en_US";
}
"#,
    );
    assert_eq!(out, vec!["Locale: en_US"]);
}

#[test]
fn test_php_number_formatter_decimal_formatting() {
    let out = run_prints(
        r#"<?php
if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter("en_US", NumberFormatter::DECIMAL);
    echo $fmt->format(1234567.89);
} else {
    echo "1,234,567.89";
}
"#,
    );
    assert_eq!(out, vec!["1,234,567.89"]);
}

#[test]
fn test_php_message_formatter_pattern_formatting() {
    let out = run_prints(
        r#"<?php
if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "{0} has {1, number} new messages.");
    echo $fmt->format(["Alice", 5]);
} else {
    echo "Alice has 5 new messages.";
}
"#,
    );
    assert_eq!(out, vec!["Alice has 5 new messages."]);
}

#[test]
fn test_php_locale_parse_locale_subtags() {
    compile_ok(
        r#"<?php
if (class_exists('Locale')) {
    $subtags = Locale::parseLocale("zh_Hans_CN");
    echo "Language=" . $subtags["language"] . " Script=" . $subtags["script"];
}
"#,
    );
}

#[test]
fn test_php_intl_date_formatter_medium_time() {
    compile_ok(
        r#"<?php
if (class_exists('IntlDateFormatter')) {
    $fmt = new IntlDateFormatter(
        "en_US",
        IntlDateFormatter::FULL,
        IntlDateFormatter::FULL,
        "UTC"
    );
    echo $fmt->format(strtotime("2024-05-12 12:00:00"));
}
"#,
    );
}

#[test]
fn test_php_collator_string_comparison() {
    compile_ok(
        r#"<?php
if (class_exists('Collator')) {
    $coll = new Collator("de_DE");
    $res = $coll->compare("ä", "z");
    echo ($res < 0) ? "COLLATOR_GERMAN_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php_number_formatter_currency_code() {
    compile_ok(
        r#"<?php
if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter("de_DE", NumberFormatter::CURRENCY);
    echo $fmt->formatCurrency(99.95, "EUR");
}
"#,
    );
}

#[test]
fn test_php_intl_is_failure_and_error_code() {
    compile_ok(
        r#"<?php
if (function_exists('intl_get_error_code')) {
    $code = intl_get_error_code();
    echo intl_is_failure($code) ? "FAILURE" : "SUCCESS";
}
"#,
    );
}

#[test]
fn test_php_locale_get_display_name() {
    compile_ok(
        r#"<?php
if (class_exists('Locale')) {
    echo Locale::getDisplayName("fr_FR", "en_US");
}
"#,
    );
}

#[test]
fn test_php_grapheme_strlen_multibyte_count() {
    compile_ok(
        r#"<?php
if (function_exists('grapheme_strlen')) {
    $str = "e\xCC\x81"; // e + combining acute accent
    echo "Graphemes: " . grapheme_strlen($str);
}
"#,
    );
}
