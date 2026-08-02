<?php
// vybe-test: php/intl/number_formatter_scientific
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::SCIENTIFIC);
echo $fmt->format(123456789.0);
