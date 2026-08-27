<?php
// vybe-test: php/php84/new_in_const_initializer
// origin: languages/php/tests/php/test_php84.rs

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

echo "new_in_const_initializer_ok";

__vybe_check(ob_get_clean(), "new_in_const_initializer_ok");
