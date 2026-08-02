<?php
// vybe-test: php/array_builtins_extended/ksort_ascending_by_key
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["banana" => 2, "apple" => 1, "cherry" => 3];
ksort($a);
echo implode(",", array_keys($a));
