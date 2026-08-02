<?php
// vybe-test: php/intl/number_formatter_currency_eur
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('de_DE', NumberFormatter::CURRENCY);
echo str_contains($fmt->formatCurrency(1234.56, 'EUR'), '1.234') ? 'de format' : 'other format';
