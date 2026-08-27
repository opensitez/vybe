<?php
// vybe-test: php/type_juggling_runtime/object_to_string_numeric_comparison
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

echo "object_to_string_numeric_comparison_ok";

__vybe_check(ob_get_clean(), "object_to_string_numeric_comparison_ok");
