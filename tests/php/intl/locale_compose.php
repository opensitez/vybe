<?php
// vybe-test: php/intl/locale_compose
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Locale')) { echo 'skipped'; return; }
$locale = Locale::composeLocale([
    'language' => 'en',
    'region'   => 'US',
]);
echo $locale;
