<?php
// vybe-test: php/closures/closure_in_usort
// origin: languages/php/tests/php/test_closures.rs

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

echo "closure_in_usort_ok";

__vybe_check(ob_get_clean(), "closure_in_usort_ok");
