<?php
// vybe-test: php/string_formatting/wordwrap_basic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$text = "The quick brown fox jumped over the lazy dog";
echo wordwrap($text, 15, "\n", false);
echo "\n";
