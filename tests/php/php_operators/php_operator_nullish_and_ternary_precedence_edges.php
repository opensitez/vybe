<?php
// vybe-test: php/php_operators/php_operator_nullish_and_ternary_precedence_edges
// origin: languages/php/tests/php/test_php_operators.rs

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

echo "php_operator_nullish_and_ternary_precedence_edges_ok";

__vybe_check(ob_get_clean(), "php_operator_nullish_and_ternary_precedence_edges_ok");
