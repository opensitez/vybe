<?php
// vybe-test: php/intl/locale_parse
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('Locale')) { echo 'skipped'; return; }
$subtags = Locale::parseLocale('zh_Hans_CN');
echo isset($subtags['language']) ? $subtags['language'] : 'no lang';
echo isset($subtags['region'])   ? ':' . $subtags['region'] : ':no region';
