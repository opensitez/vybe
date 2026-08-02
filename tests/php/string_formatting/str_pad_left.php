<?php
// vybe-test: php/string_formatting/str_pad_left
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo str_pad("42", 6, "0", STR_PAD_LEFT);
echo "\n";
echo str_pad("x", 5, ".", STR_PAD_LEFT);
echo "\n";
