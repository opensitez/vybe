<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_throw_expression_in_match_arm
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs

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

echo "test_php_throw_expression_in_match_arm_ok";

__vybe_check(ob_get_clean(), "test_php_throw_expression_in_match_arm_ok");
