<?php
// vybe-test: php/array_map_multiple/array_map_null_single_array_wraps_in_array
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

echo "array_map_null_single_array_wraps_in_array_ok";

__vybe_check(ob_get_clean(), "array_map_null_single_array_wraps_in_array_ok");
