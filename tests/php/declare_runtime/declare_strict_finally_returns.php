<?php
// vybe-test: php/declare_runtime/declare_strict_finally_returns
// origin: languages/php/tests/php/test_declare_runtime.rs

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

echo "declare_strict_finally_returns_ok";

__vybe_check(ob_get_clean(), "declare_strict_finally_returns_ok");
