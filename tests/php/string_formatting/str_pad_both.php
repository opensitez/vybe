<?php
// vybe-test: php/string_formatting/str_pad_both
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo str_pad("hi", 8, "-", STR_PAD_BOTH) . "|";
echo "\n";
echo str_pad("a", 5, "*", STR_PAD_BOTH)  . "|";
echo "\n";
