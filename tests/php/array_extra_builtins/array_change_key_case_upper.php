<?php
// vybe-test: php/array_extra_builtins/array_change_key_case_upper
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["first" => 1, "Second" => 2, "THIRD" => 3];
$upper = array_change_key_case($a, CASE_UPPER);
echo array_key_exists("FIRST", $upper) ? "yes" : "no";
echo array_key_exists("SECOND", $upper) ? "yes" : "no";
echo array_key_exists("THIRD", $upper) ? "yes" : "no";
