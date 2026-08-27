<?php
// vybe-test: php/generator_send_runtime/generator_empty_yields_nothing
// origin: languages/php/tests/php/test_generator_send_runtime.rs

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

echo "generator_empty_yields_nothing_ok";

__vybe_check(ob_get_clean(), "generator_empty_yields_nothing_ok");
