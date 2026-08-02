<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_type_check_builtin
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$ints = [10, 20, 30];
$res = function_exists('array_all')
    ? array_all($ints, "is_int")
    : true;
echo $res ? "ALL_INTS_OK" : "FAIL";
