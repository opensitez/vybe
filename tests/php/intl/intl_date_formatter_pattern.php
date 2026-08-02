<?php
// vybe-test: php/intl/intl_date_formatter_pattern
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::NONE, IntlDateFormatter::NONE,
    'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd');
$result = $fmt->format(mktime(0, 0, 0, 6, 15, 2024));
echo $result;
