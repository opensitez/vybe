<?php
// vybe-test: php/intl/number_formatter_attributes
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
$fmt->setAttribute(NumberFormatter::MIN_FRACTION_DIGITS, 2);
$fmt->setAttribute(NumberFormatter::MAX_FRACTION_DIGITS, 4);
echo $fmt->format(3.14159);
