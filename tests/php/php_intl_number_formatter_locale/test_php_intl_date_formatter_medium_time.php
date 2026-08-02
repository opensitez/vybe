<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_intl_date_formatter_medium_time
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (class_exists('IntlDateFormatter')) {
    $fmt = new IntlDateFormatter(
        "en_US",
        IntlDateFormatter::FULL,
        IntlDateFormatter::FULL,
        "UTC"
    );
    echo $fmt->format(strtotime("2024-05-12 12:00:00"));
}
