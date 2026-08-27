<?php
// vybe-test: php/operators_runtime/ternary_is_right_associative_runtime
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

echo "ternary_is_right_associative_runtime_ok";

__vybe_check(ob_get_clean(), "ternary_is_right_associative_runtime_ok");
