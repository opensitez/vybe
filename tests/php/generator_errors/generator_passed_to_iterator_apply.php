<?php
// vybe-test: php/generator_errors/generator_passed_to_iterator_apply
// origin: languages/php/tests/php/test_generator_errors.rs

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

echo "generator_passed_to_iterator_apply_ok";

__vybe_check(ob_get_clean(), "generator_passed_to_iterator_apply_ok");
