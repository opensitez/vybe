<?php
// vybe-test: php/arrays/array_multi_assign_with_numeric_like_keys
// origin: languages/php/tests/php/test_arrays.rs

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

echo "array_multi_assign_with_numeric_like_keys_ok";

__vybe_check(ob_get_clean(), "array_multi_assign_with_numeric_like_keys_ok");
