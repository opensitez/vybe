<?php
// vybe-test: php/intl/intl_date_formatter_basic
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::LONG, IntlDateFormatter::NONE);
$result = $fmt->format(mktime(0, 0, 0, 1, 15, 2024));
echo strlen($result) > 0 ? 'formatted' : 'empty';
