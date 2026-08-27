<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_ternary_nested_with_parentheses_is_right_associative
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs

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

echo "test_php_ternary_nested_with_parentheses_is_right_associative_ok";

__vybe_check(ob_get_clean(), "test_php_ternary_nested_with_parentheses_is_right_associative_ok");
