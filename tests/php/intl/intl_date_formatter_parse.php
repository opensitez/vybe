<?php
// vybe-test: php/intl/intl_date_formatter_parse
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::NONE, IntlDateFormatter::NONE,
    'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd');
$ts = $fmt->parse('2024-06-15');
echo $ts !== false ? 'parsed' : 'failed';
