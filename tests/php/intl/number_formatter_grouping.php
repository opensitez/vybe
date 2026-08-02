<?php
// vybe-test: php/intl/number_formatter_grouping
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('fr_FR', NumberFormatter::DECIMAL);
$formatted = $fmt->format(1000000);
echo strlen($formatted) > 0 ? 'formatted' : 'empty';
