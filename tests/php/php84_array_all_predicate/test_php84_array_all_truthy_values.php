<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_truthy_values
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$truthies = [1, "text", true, [1]];
$res = function_exists('array_all')
    ? array_all($truthies, fn($v) => (bool)$v)
    : true;
echo $res ? "ALL_TRUTHY_OK" : "FAIL";
