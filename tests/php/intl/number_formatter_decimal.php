<?php
// vybe-test: php/intl/number_formatter_decimal
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'intl not available'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
echo $fmt->format(1234567.89);
