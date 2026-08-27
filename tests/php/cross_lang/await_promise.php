<?php
// vybe-test: php/cross_lang/await_promise
// origin: languages/php/tests/php/test_cross_lang.rs

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

echo "await_promise_ok";

__vybe_check(ob_get_clean(), "await_promise_ok");
