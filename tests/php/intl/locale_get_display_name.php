<?php
// vybe-test: php/intl/locale_get_display_name
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Locale')) { echo 'skipped'; return; }
$name = Locale::getDisplayName('en_US', 'en');
echo strlen($name) > 0 ? 'has name' : 'empty';
