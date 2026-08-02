<?php
// vybe-test: php/string_formatting/sprintf_char
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%c", 65);   // A
echo sprintf("%c", 97);   // a
echo sprintf("%c%c%c", 72, 105, 33);  // Hi!
