<?php
// vybe-test: php/intl/locale_get_default
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Locale')) { echo 'skipped'; return; }
$locale = Locale::getDefault();
echo is_string($locale) ? 'is string' : 'not string';
