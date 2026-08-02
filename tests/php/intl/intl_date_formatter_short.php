<?php
// vybe-test: php/intl/intl_date_formatter_short
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('en_US', IntlDateFormatter::SHORT, IntlDateFormatter::SHORT);
$result = $fmt->format(mktime(14, 30, 0, 3, 15, 2024));
echo strlen($result) > 0 ? 'short formatted' : 'empty';
