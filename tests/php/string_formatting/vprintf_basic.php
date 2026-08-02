<?php
// vybe-test: php/string_formatting/vprintf_basic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$args = ["Bob", 25];
$written = vprintf("Name: %s, Age: %d\n", $args);
echo $written > 0 ? 'wrote bytes' : 'nothing';
echo "\n";
