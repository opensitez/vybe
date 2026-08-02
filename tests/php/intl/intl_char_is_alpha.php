<?php
// vybe-test: php/intl/intl_char_is_alpha
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlChar')) { echo 'skipped'; return; }
echo IntlChar::isalpha('A') ? 'is alpha' : 'not alpha';
echo IntlChar::isalpha('1') ? 'is alpha' : ':not alpha';
