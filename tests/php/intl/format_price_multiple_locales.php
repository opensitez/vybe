<?php
// vybe-test: php/intl/format_price_multiple_locales
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$amount = 1234567.89;
$locales = ['en_US' => 'USD', 'de_DE' => 'EUR', 'ja_JP' => 'JPY'];
foreach ($locales as $locale => $currency) {
    $fmt = new NumberFormatter($locale, NumberFormatter::CURRENCY);
    $formatted = $fmt->formatCurrency($amount, $currency);
    echo $locale . ': ' . $formatted . "\n";
}
echo 'done';
