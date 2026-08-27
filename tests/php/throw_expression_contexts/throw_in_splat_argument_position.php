<?php
// vybe-test: php/throw_expression_contexts/throw_in_splat_argument_position
// origin: languages/php/tests/php/test_throw_expression_contexts.rs

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

echo "throw_in_splat_argument_position_ok";

__vybe_check(ob_get_clean(), "throw_in_splat_argument_position_ok");
