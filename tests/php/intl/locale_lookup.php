<?php
// vybe-test: php/intl/locale_lookup
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Locale')) { echo 'skipped'; return; }
$match = Locale::lookup(['fr_FR', 'fr', 'en_US', 'en'], 'fr_CA', true, 'en');
echo $match;
