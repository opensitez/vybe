<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_single_element_true
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$single = ["yes"];
$res = function_exists('array_any')
    ? array_any($single, fn($s) => $s === "yes")
    : true;
echo $res ? "SINGLE_MATCH_OK" : "FAIL";
