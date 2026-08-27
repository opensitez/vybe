<?php
// vybe-test: php/operators_runtime/assignment_and_ternary_precedence_runtime
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

echo "assignment_and_ternary_precedence_runtime_ok";

__vybe_check(ob_get_clean(), "assignment_and_ternary_precedence_runtime_ok");
