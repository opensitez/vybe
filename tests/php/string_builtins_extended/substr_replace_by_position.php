<?php
// vybe-test: php/string_builtins_extended/substr_replace_by_position
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo substr_replace("Hello World", "PHP", 6, 5);
echo substr_replace("abcdefgh", "XYZ", 2, 3);
echo substr_replace("insert here", ">>", 6, 0);
