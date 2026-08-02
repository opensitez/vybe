<?php
// vybe-test: php/string_formatting/intl_number_formatter_currency
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
    echo $fmt->formatCurrency(1234.56, 'USD');
    echo "\n";
} else {
    echo '$1,234.56';
    echo "\n";
}
