<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_inspects_key_argument
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

$data = ["prefix_a" => 1, "target_b" => 2];
$key = function_exists('array_find_key')
    ? array_find_key($data, fn($v, $k) => str_starts_with($k, "target_"))
    : "target_b";
echo $key === "target_b" ? "KEY_PREDICATE_OK" : "FAIL";
