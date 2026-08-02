<?php
// vybe-test: php/string_formatting/str_pad_right
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo str_pad("hello", 10) . "|";
echo "\n";
echo str_pad("hi", 8, "-") . "|";
echo "\n";
