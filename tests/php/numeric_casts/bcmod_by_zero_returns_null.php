<?php
// vybe-test: php/numeric_casts/bcmod_by_zero_returns_null
// origin: languages/php/tests/php/test_numeric_casts.rs

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

echo "bcmod_by_zero_returns_null_ok";

__vybe_check(ob_get_clean(), "bcmod_by_zero_returns_null_ok");
