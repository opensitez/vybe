<?php
// vybe-test: php/string_builtins_extended/vsprintf_array_args
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$result = vsprintf("Name: %s, Score: %d, Ratio: %.2f", ["Alice", 95, 0.987]);
echo $result;
