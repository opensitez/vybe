<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_string_lengths
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$words = ["hello", "world", "php84"];
$res = function_exists('array_all')
    ? array_all($words, fn($w) => strlen($w) === 5)
    : true;
echo $res ? "ALL_LEN_5_OK" : "FAIL";
