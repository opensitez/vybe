<?php
// vybe-test: php/output_functions/printf_multiple_specifiers
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$written = printf("Name: %s, Age: %d, Score: %.2f\n", "Alice", 30, 98.5);
echo $written > 0 ? 'wrote bytes' : 'nothing written';
