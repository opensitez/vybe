<?php
// vybe-test: php/string_formatting/fprintf_to_stdout
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$written = fprintf(STDOUT, "Value: %d\n", 42);
echo $written > 0 ? 'wrote' : 'nothing';
echo "\n";
