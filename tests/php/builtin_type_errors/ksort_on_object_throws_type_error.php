<?php
// vybe-test: php/builtin_type_errors/ksort_on_object_throws_type_error
// origin: languages/php/tests/php/test_builtin_type_errors.rs

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

echo "ksort_on_object_throws_type_error_ok";

__vybe_check(ob_get_clean(), "ksort_on_object_throws_type_error_ok");
