<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_short_circuits_on_first_true
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$calls = 0;
$arr = [10, 20, 30];
if (function_exists('array_any')) {
    array_any($arr, function($v) use (&$calls) {
        $calls++;
        return $v >= 10;
    });
    echo $calls === 1 ? "SHORT_CIRCUIT_OK" : "FAIL";
} else {
    echo "SHORT_CIRCUIT_OK";
}
