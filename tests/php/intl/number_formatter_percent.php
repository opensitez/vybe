<?php
// vybe-test: php/intl/number_formatter_percent
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::PERCENT);
echo $fmt->format(0.75);
