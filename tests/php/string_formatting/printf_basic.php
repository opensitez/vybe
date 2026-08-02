<?php
// vybe-test: php/string_formatting/printf_basic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$written = printf("Name: %s, Age: %d\n", "Alice", 30);
echo $written > 0 ? 'wrote bytes' : 'nothing written';
echo "\n";
