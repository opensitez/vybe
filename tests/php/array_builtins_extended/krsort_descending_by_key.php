<?php
// vybe-test: php/array_builtins_extended/krsort_descending_by_key
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["alpha" => 10, "gamma" => 30, "beta" => 20];
krsort($a);
echo implode(",", array_keys($a));
