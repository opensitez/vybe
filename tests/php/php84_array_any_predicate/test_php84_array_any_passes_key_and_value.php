<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_passes_key_and_value
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$map = ["role" => "admin", "active" => true];
$res = function_exists('array_any')
    ? array_any($map, fn($v, $k) => $k === "role" && $v === "admin")
    : true;
echo $res ? "KEY_VAL_PASS_OK" : "FAIL";
