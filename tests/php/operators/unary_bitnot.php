<?php
// vybe-test: php/operators/unary_bitnot
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

echo "unary_bitnot_ok";

__vybe_check(ob_get_clean(), "unary_bitnot_ok");
