<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_negative_numbers
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$negatives = [-1, -5, -10];
$res = function_exists('array_all')
    ? array_all($negatives, fn($n) => $n < 0)
    : true;
echo $res ? "ALL_NEGATIVES_OK" : "FAIL";
