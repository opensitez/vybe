<?php
// vybe-test: php/object_comparison/array_unique_with_equal_objects_keeps_first
// origin: languages/php/tests/php/test_object_comparison.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "array_unique_with_equal_objects_keeps_first_ok";

__vybe_check(ob_get_clean(), "array_unique_with_equal_objects_keeps_first_ok");
