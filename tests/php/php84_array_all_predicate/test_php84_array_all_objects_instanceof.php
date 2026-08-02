<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_objects_instanceof
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs
// vybe-test-mode: compile

$objects = [new stdClass(), new stdClass()];
$res = function_exists('array_all')
    ? array_all($objects, fn($o) => $o instanceof stdClass)
    : true;
echo $res ? "ALL_STDCLASS_OK" : "FAIL";
