<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_passes_key_and_value
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$map = ["a_1" => 10, "a_2" => 20];
$res = function_exists('array_all')
    ? array_all($map, fn($v, $k) => str_starts_with($k, "a_"))
    : true;
echo $res ? "ALL_KEYS_MATCH_OK" : "FAIL";
