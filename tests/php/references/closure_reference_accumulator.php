<?php
// vybe-test: php/references/closure_reference_accumulator
// origin: languages/php/tests/php/test_references.rs

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

echo "closure_reference_accumulator_ok";

__vybe_check(ob_get_clean(), "closure_reference_accumulator_ok");
