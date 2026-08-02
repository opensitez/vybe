<?php
// vybe-test: php/string_formatting/wordwrap_cut_long_words
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$text = "A verylongwordthatcannotfit in normal wrapping";
echo wordwrap($text, 10, "\n", true);
echo "\n";
