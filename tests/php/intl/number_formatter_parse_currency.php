<?php
// vybe-test: php/intl/number_formatter_parse_currency
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
$currency = '';
$amount = $fmt->parseCurrency('$1,234.56', $currency);
echo round($amount, 2) . ':' . $currency;
