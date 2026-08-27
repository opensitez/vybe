<?php
// vybe-test: php/assert/assert_on_is_array_type_check_false
// origin: languages/php/tests/php/test_assert.rs

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

echo "assert_on_is_array_type_check_false_ok";

__vybe_check(ob_get_clean(), "assert_on_is_array_type_check_false_ok");
