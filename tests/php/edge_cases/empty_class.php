<?php
// vybe-test: php/edge_cases/empty_class
// origin: languages/php/tests/php/test_edge_cases.rs

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

echo "empty_class_ok";

__vybe_check(ob_get_clean(), "empty_class_ok");
