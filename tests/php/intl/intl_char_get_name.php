<?php
// vybe-test: php/intl/intl_char_get_name
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('IntlChar')) { echo 'skipped'; return; }
$name = IntlChar::charName('A');
echo strlen($name) > 0 ? 'has name' : 'empty';
