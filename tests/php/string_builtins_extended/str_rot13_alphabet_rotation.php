<?php
// vybe-test: php/string_builtins_extended/str_rot13_alphabet_rotation
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$msg = "Hello World 123";
$rotated = str_rot13($msg);
echo $rotated;
echo str_rot13($rotated);
