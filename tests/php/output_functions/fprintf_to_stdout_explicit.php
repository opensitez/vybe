<?php
// vybe-test: php/output_functions/fprintf_to_stdout_explicit
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$written = fprintf(STDOUT, "Item: %s costs $%.2f\n", "widget", 4.99);
echo $written > 0 ? 'wrote' : 'nothing';
