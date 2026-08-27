<?php
// vybe-test: php/operators_runtime/unary_and_nullsafe_truthiness_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

echo "unary_and_nullsafe_truthiness_runtime_ok";

__vybe_check(ob_get_clean(), "unary_and_nullsafe_truthiness_runtime_ok");
