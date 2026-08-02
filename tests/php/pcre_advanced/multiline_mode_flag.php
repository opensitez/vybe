<?php
// vybe-test: php/pcre_advanced/multiline_mode_flag
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$text = "first line\nsecond line\nthird line";
preg_match_all('/^\w+/m', $text, $m);
echo implode(',', $m[0]);
