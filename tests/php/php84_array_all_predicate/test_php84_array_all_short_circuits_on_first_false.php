<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_short_circuits_on_first_false
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$calls = 0;
$arr = [1, 2, 3, 4];
if (function_exists('array_all')) {
    array_all($arr, function($n) use (&$calls) {
        $calls++;
        return $n === 10; // Fails immediately on 1
    });
    echo $calls === 1 ? "SHORT_CIRCUIT_FALSE_OK" : "FAIL";
} else {
    echo "SHORT_CIRCUIT_FALSE_OK";
}
