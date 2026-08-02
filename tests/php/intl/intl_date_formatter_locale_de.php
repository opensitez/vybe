<?php
// vybe-test: php/intl/intl_date_formatter_locale_de
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlDateFormatter')) { echo 'skipped'; return; }
$fmt = new IntlDateFormatter('de_DE', IntlDateFormatter::FULL, IntlDateFormatter::NONE);
$result = $fmt->format(mktime(0, 0, 0, 12, 25, 2024));
echo strlen($result) > 0 ? 'de formatted' : 'empty';
