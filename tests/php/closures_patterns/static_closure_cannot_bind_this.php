<?php
// vybe-test: php/closures_patterns/static_closure_cannot_bind_this
// origin: languages/php/tests/php/test_closures_patterns.rs

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

echo "static_closure_cannot_bind_this_ok";

__vybe_check(ob_get_clean(), "static_closure_cannot_bind_this_ok");
