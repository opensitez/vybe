<?php
// vybe-test: php/array_extra_builtins/array_rand_single_random_key
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["alpha" => 1, "beta" => 2, "gamma" => 3];
$key = array_rand($a);
echo array_key_exists($key, $a) ? "found" : "missing";
