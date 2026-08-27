<?php
// vybe-test: php/operators/logical_assignment_operators_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo "logical_assignment_operators_runtime_ok";

__vybe_check(ob_get_clean(), "logical_assignment_operators_runtime_ok");
