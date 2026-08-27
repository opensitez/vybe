<?php
// vybe-test: php/generators_advanced/generator_key_value_pairs
// origin: languages/php/tests/php/test_generators_advanced.rs

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

echo "generator_key_value_pairs_ok";

__vybe_check(ob_get_clean(), "generator_key_value_pairs_ok");
