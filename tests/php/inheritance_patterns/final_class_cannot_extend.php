<?php
// vybe-test: php/inheritance_patterns/final_class_cannot_extend
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

echo "final_class_cannot_extend_ok";

__vybe_check(ob_get_clean(), "final_class_cannot_extend_ok");
