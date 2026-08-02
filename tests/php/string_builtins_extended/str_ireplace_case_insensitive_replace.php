<?php
// vybe-test: php/string_builtins_extended/str_ireplace_case_insensitive_replace
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$result = str_ireplace("HELLO", "Hi", "Hello World HELLO hello");
echo $result;
