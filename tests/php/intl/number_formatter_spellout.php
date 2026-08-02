<?php
// vybe-test: php/intl/number_formatter_spellout
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('NumberFormatter')) { echo 'skipped'; return; }
$fmt = new NumberFormatter('en_US', NumberFormatter::SPELLOUT);
$spelled = $fmt->format(42);
echo strlen($spelled) > 0 ? 'spelled' : 'empty';
