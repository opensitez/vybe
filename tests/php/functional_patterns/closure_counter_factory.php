<?php
// vybe-test: php/functional_patterns/closure_counter_factory
// origin: languages/php/tests/php/test_functional_patterns.rs

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

echo "closure_counter_factory_ok";

__vybe_check(ob_get_clean(), "closure_counter_factory_ok");
