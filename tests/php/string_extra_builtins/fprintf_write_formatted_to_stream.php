<?php
// vybe-test: php/string_extra_builtins/fprintf_write_formatted_to_stream
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$written = fprintf(STDOUT, "Name: %s, Age: %d, Score: %.2f\n", "Alice", 30, 98.5);
echo is_int($written) ? "int" : "not-int";
