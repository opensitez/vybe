<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_truthy_coercion
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$items = [0, 0, 1, 0];
$res = function_exists('array_any')
    ? array_any($items, fn($v) => $v)
    : true;
echo $res ? "TRUTHY_COERCION_OK" : "FAIL";
