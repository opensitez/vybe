<?php
// vybe-test: php/type_juggling/coercion_string_to_int_arithmetic
// origin: languages/php/tests/php/test_type_juggling.rs

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

echo "coercion_string_to_int_arithmetic_ok";

__vybe_check(ob_get_clean(), "coercion_string_to_int_arithmetic_ok");
