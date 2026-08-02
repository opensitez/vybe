<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_number_formatter_currency_code
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (class_exists('NumberFormatter')) {
    $fmt = new NumberFormatter("de_DE", NumberFormatter::CURRENCY);
    echo $fmt->formatCurrency(99.95, "EUR");
}
