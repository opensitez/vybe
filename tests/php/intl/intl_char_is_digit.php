<?php
// vybe-test: php/intl/intl_char_is_digit
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::isdigit('5') ? 'is digit' : 'not digit';
echo IntlChar::isdigit('a') ? 'is digit' : ':not digit';
