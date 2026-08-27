<?php
// vybe-test: php/abstract_final_patterns/final_method_cannot_be_overridden
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

echo "final_method_cannot_be_overridden_ok";

__vybe_check(ob_get_clean(), "final_method_cannot_be_overridden_ok");
