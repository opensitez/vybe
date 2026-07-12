use super::helpers::compile_ok;

// ── NumberFormatter ───────────────────────────────────────────

#[test]
fn number_formatter_decimal() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'intl not available'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
echo $fmt->format(1234567.89);
"#,
    );
}

#[test]
fn number_formatter_currency() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
echo $fmt->formatCurrency(1234.56, 'USD');
"#,
    );
}

#[test]
fn number_formatter_currency_eur() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('de_DE', NumberFormatter::CURRENCY);
echo str_contains($fmt->formatCurrency(1234.56, 'EUR'), '1.234') ? 'de format' : 'other format';
"#,
    );
}

#[test]
fn number_formatter_percent() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::PERCENT);
echo $fmt->format(0.75);
"#,
    );
}

#[test]
fn number_formatter_scientific() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::SCIENTIFIC);
echo $fmt->format(123456789.0);
"#,
    );
}

#[test]
fn number_formatter_spellout() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::SPELLOUT);
$spelled = $fmt->format(42);
echo strlen($spelled) > 0 ? 'spelled' : 'empty';
"#,
    );
}

#[test]
fn number_formatter_parse() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
$num = $fmt->parse('1,234,567.89');
echo round($num, 2);
"#,
    );
}

#[test]
fn number_formatter_parse_currency() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
$currency = '';
$amount = $fmt->parseCurrency('$1,234.56', $currency);
echo round($amount, 2) . ':' . $currency;
"#,
    );
}

#[test]
fn number_formatter_attributes() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
$fmt->setAttribute(NumberFormatter::MIN_FRACTION_DIGITS, 2);
$fmt->setAttribute(NumberFormatter::MAX_FRACTION_DIGITS, 4);
echo $fmt->format(3.14159);
"#,
    );
}

#[test]
fn number_formatter_grouping() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('fr_FR', NumberFormatter::DECIMAL);
$formatted = $fmt->format(1000000);
echo strlen($formatted) > 0 ? 'formatted' : 'empty';
"#,
    );
}

// ── IntlDateFormatter ─────────────────────────────────────────

#[test]
fn intl_date_formatter_basic() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::LONG, IntlDateFormatter::NONE);
$result = $fmt->format(mktime(0, 0, 0, 1, 15, 2024));
echo strlen($result) > 0 ? 'formatted' : 'empty';
"#,
    );
}

#[test]
fn intl_date_formatter_short() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::SHORT, IntlDateFormatter::SHORT);
$result = $fmt->format(mktime(14, 30, 0, 3, 15, 2024));
echo strlen($result) > 0 ? 'short formatted' : 'empty';
"#,
    );
}

#[test]
fn intl_date_formatter_pattern() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::NONE, IntlDateFormatter::NONE,
    'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd');
$result = $fmt->format(mktime(0, 0, 0, 6, 15, 2024));
echo $result;
"#,
    );
}

#[test]
fn intl_date_formatter_parse() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::NONE, IntlDateFormatter::NONE,
    'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd');
$ts = $fmt->parse('2024-06-15');
echo $ts !== false ? 'parsed' : 'failed';
"#,
    );
}

#[test]
fn intl_date_formatter_locale_de() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('de_DE', IntlDateFormatter::FULL, IntlDateFormatter::NONE);
$result = $fmt->format(mktime(0, 0, 0, 12, 25, 2024));
echo strlen($result) > 0 ? 'de formatted' : 'empty';
"#,
    );
}

// ── Collator ─────────────────────────────────────────────────

#[test]
fn collator_basic_compare() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
echo $coll->compare('apple', 'Banana') < 0 ? 'apple < Banana' : 'apple >= Banana';
"#,
    );
}

#[test]
fn collator_sort() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$words = ['Banana', 'apple', 'Cherry', 'date'];
$coll->sort($words);
echo implode(',', $words);
"#,
    );
}

#[test]
fn collator_sort_locale_de() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('de_DE');
$words = ['Österreich', 'Angola', 'Zürich', 'Belgien'];
$coll->sort($words);
echo implode(',', $words);
"#,
    );
}

#[test]
fn collator_strength() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$coll->setStrength(Collator::PRIMARY);
echo $coll->compare('Cafe', 'café') === 0 ? 'equal (primary)' : 'different';
"#,
    );
}

#[test]
fn collator_get_sort_key() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$coll = new Collator('en_US');
$key1 = $coll->getSortKey('apple');
$key2 = $coll->getSortKey('banana');
echo ($key1 < $key2) ? 'apple < banana' : 'apple >= banana';
"#,
    );
}

// ── MessageFormatter ──────────────────────────────────────────

#[test]
fn message_formatter_basic() {
    compile_ok(
        r#"<?php
if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$fmt = new MessageFormatter('en_US', 'Hello, {0}!');
echo $fmt->format(['World']);
"#,
    );
}

#[test]
fn message_formatter_plural() {
    compile_ok(
        r#"<?php
if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$pattern = '{0, plural, =0{no items} one{# item} other{# items}}';
$fmt = new MessageFormatter('en_US', $pattern);
echo $fmt->format([0]) . ':';
echo $fmt->format([1]) . ':';
echo $fmt->format([5]);
"#,
    );
}

#[test]
fn message_formatter_static() {
    compile_ok(
        r#"<?php
if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$result = MessageFormatter::formatMessage('en_US', 'Value: {0, number}', [1234567.89]);
echo strlen($result) > 0 ? 'formatted' : 'empty';
"#,
    );
}

#[test]
fn message_formatter_date() {
    compile_ok(
        r#"<?php
if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$fmt = new MessageFormatter('en_US', 'Date: {0, date, short}');
$result = $fmt->format([time()]);
echo strlen($result) > 0 ? 'date formatted' : 'empty';
"#,
    );
}

// ── Locale ───────────────────────────────────────────────────

#[test]
fn locale_get_default() {
    compile_ok(
        r#"<?php
if (!class_exists('Locale')) { echo 'skipped'; return; }
$locale = Locale::getDefault();
echo is_string($locale) ? 'is string' : 'not string';
"#,
    );
}

#[test]
fn locale_parse() {
    compile_ok(
        r#"<?php
if (!class_exists('Locale')) { echo 'skipped'; return; }
$subtags = Locale::parseLocale('zh_Hans_CN');
echo isset($subtags['language']) ? $subtags['language'] : 'no lang';
echo isset($subtags['region'])   ? ':' . $subtags['region'] : ':no region';
"#,
    );
}

#[test]
fn locale_compose() {
    compile_ok(
        r#"<?php
if (!class_exists('Locale')) { echo 'skipped'; return; }
$locale = Locale::composeLocale([
    'language' => 'en',
    'region'   => 'US',
]);
echo $locale;
"#,
    );
}

#[test]
fn locale_lookup() {
    compile_ok(
        r#"<?php
if (!class_exists('Locale')) { echo 'skipped'; return; }
$match = Locale::lookup(['fr_FR', 'fr', 'en_US', 'en'], 'fr_CA', true, 'en');
echo $match;
"#,
    );
}

#[test]
fn locale_get_display_name() {
    compile_ok(
        r#"<?php
if (!class_exists('Locale')) { echo 'skipped'; return; }
$name = Locale::getDisplayName('en_US', 'en');
echo strlen($name) > 0 ? 'has name' : 'empty';
"#,
    );
}

// ── IntlChar (PHP 7.0+) ───────────────────────────────────────

#[test]
fn intl_char_is_alpha() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::isalpha('A') ? 'is alpha' : 'not alpha';
echo IntlChar::isalpha('1') ? 'is alpha' : ':not alpha';
"#,
    );
}

#[test]
fn intl_char_is_digit() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::isdigit('5') ? 'is digit' : 'not digit';
echo IntlChar::isdigit('a') ? 'is digit' : ':not digit';
"#,
    );
}

#[test]
fn intl_char_to_upper_lower() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::toupper('a');
echo IntlChar::tolower('Z');
"#,
    );
}

#[test]
fn intl_char_get_name() {
    compile_ok(
        r#"<?php
if (!class_exists('IntlChar')) { echo 'skipped'; return; }
$name = IntlChar::charName('A');
echo strlen($name) > 0 ? 'has name' : 'empty';
"#,
    );
}

// ── Practical intl patterns ───────────────────────────────────

#[test]
fn format_price_multiple_locales() {
    compile_ok(
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$amount = 1234567.89;
$locales = ['en_US' => 'USD', 'de_DE' => 'EUR', 'ja_JP' => 'JPY'];
foreach ($locales as $locale => $currency) {
    $fmt = new NumberFormatter($locale, NumberFormatter::CURRENCY);
    $formatted = $fmt->formatCurrency($amount, $currency);
    echo $locale . ': ' . $formatted . "\n";
}
echo 'done';
"#,
    );
}

#[test]
fn sort_names_locale_aware() {
    compile_ok(
        r#"<?php
if (!class_exists('Collator')) { echo 'skipped'; return; }
$names = ['Müller', 'Maier', 'Möller', 'Meyer'];
$coll = new Collator('de_DE');
$coll->sort($names);
echo implode(',', $names);
"#,
    );
}
