<?php
// vybe-test: php/intl/intl_char_to_upper_lower
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::toupper('a');
echo IntlChar::tolower('Z');
