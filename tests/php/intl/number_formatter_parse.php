<?php
// vybe-test: php/intl/number_formatter_parse
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
$num = $fmt->parse('1,234,567.89');
echo round($num, 2);
