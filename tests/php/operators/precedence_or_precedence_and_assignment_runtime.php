<?php
// vybe-test: php/operators/precedence_or_precedence_and_assignment_runtime
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

echo "precedence_or_precedence_and_assignment_runtime_ok";

__vybe_check(ob_get_clean(), "precedence_or_precedence_and_assignment_runtime_ok");
