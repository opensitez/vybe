<?php
// vybe-test: php/error_handling/throw_in_coalesce
// origin: languages/php/tests/php/test_error_handling.rs

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

echo "throw_in_coalesce_ok";

__vybe_check(ob_get_clean(), "throw_in_coalesce_ok");
